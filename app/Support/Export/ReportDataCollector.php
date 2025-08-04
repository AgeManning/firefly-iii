<?php

/**
 * ReportDataCollector.php
 * Copyright (c) 2019 james@firefly-iii.org
 *
 * This file is part of Firefly III (https://github.com/firefly-iii).
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

declare(strict_types=1);

namespace FireflyIII\Support\Export;

use Carbon\Carbon;
use FireflyIII\Http\Controllers\Report\AccountController;
use FireflyIII\Http\Controllers\Report\BalanceController;
use FireflyIII\Http\Controllers\Report\BillController;
use FireflyIII\Http\Controllers\Report\BudgetController;
use FireflyIII\Http\Controllers\Report\CategoryController;
use FireflyIII\Http\Controllers\Report\DoubleController;
use FireflyIII\Http\Controllers\Report\OperationsController;
use FireflyIII\Http\Controllers\Report\TagController;
use FireflyIII\Http\Controllers\Chart\ReportController as ChartReportController;
use FireflyIII\Http\Controllers\Chart\BudgetReportController;
use FireflyIII\Http\Controllers\Chart\CategoryReportController;
use FireflyIII\Http\Controllers\Chart\DoubleReportController;
use FireflyIII\Http\Controllers\Chart\ExpenseReportController;
use FireflyIII\Http\Controllers\Chart\TagReportController;
use FireflyIII\User;
use Illuminate\Support\Collection;
use ReflectionClass;
use FireflyIII\Repositories\Account\AccountTaskerInterface;
use FireflyIII\Support\Report\Category\CategoryReportGenerator;

/**
 * Class ReportDataCollector
 * 
 * Collects report data programmatically by calling report controllers directly
 * Bypasses the normal AJAX loading system for export purposes
 */
class ReportDataCollector
{
    private User $user;
    private string $reportType;
    private Carbon $start;
    private Carbon $end;
    private Collection $accounts;
    private Collection $budgets;
    private Collection $categories;
    private Collection $tags;
    private Collection $expenseAccounts;

    public function __construct()
    {
        $this->accounts = new Collection();
        $this->budgets = new Collection();
        $this->categories = new Collection();
        $this->tags = new Collection();
        $this->expenseAccounts = new Collection();
    }

    /**
     * Collect all report data based on report type
     */
    public function collectReportData(): array
    {
        $data = [
            'accounts' => $this->collectAccountData(),
            'balance' => $this->collectBalanceData(),
            'income' => $this->collectIncomeData(),
            'expenses' => $this->collectExpensesData(),
            'operations' => $this->collectOperationsData(),
            'charts' => $this->collectChartData(),
        ];

        // Add report-type specific data
        switch ($this->reportType) {
            case 'budget':
                $data['budgets'] = $this->collectBudgetData();
                $data['budget_charts'] = $this->collectBudgetChartData();
                break;
                
            case 'category':
                $data['categories'] = $this->collectCategoryData();
                $data['category_charts'] = $this->collectCategoryChartData();
                break;
                
            case 'tag':
                $data['tags'] = $this->collectTagData();
                $data['tag_charts'] = $this->collectTagChartData();
                break;
                
            case 'double':
                $data['double'] = $this->collectDoubleData();
                $data['double_charts'] = $this->collectDoubleChartData();
                break;
                
            default:
                // For default reports, also collect category data
                $data['categories'] = $this->collectCategoryOperationsData();
                break;
        }

        // Always collect bills data
        $data['bills'] = $this->collectBillData();

        return $data;
    }

    /**
     * Collect account data
     */
    private function collectAccountData(): array
    {
        try {
            $controller = app(AccountController::class);
            $response = $controller->accounts($this->accounts, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            if (is_string($response)) {
                return ['html' => $response];
            }
            
            if (method_exists($response, 'render')) {
                return ['html' => $response->render()];
            }
            
            return ['html' => (string)$response];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect account data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect balance data (Income vs Expenses summary)
     */
    private function collectBalanceData(): array
    {
        try {
            // Get raw data directly from AccountTasker like OperationsController does
            $tasker = app(AccountTaskerInterface::class);
            $tasker->setUser($this->user);
            
            $incomes = $tasker->getIncomeReport($this->start, $this->end, $this->accounts);
            $expenses = $tasker->getExpenseReport($this->start, $this->end, $this->accounts);
            $sums = [];
            $keys = array_unique(array_merge(array_keys($incomes['sums']), array_keys($expenses['sums'])));

            foreach ($keys as $currencyId) {
                $currencyInfo = $incomes['sums'][$currencyId] ?? $expenses['sums'][$currencyId];
                $sums[$currencyId] = [
                    'currency_id' => $currencyId,
                    'currency_name' => $currencyInfo['currency_name'],
                    'currency_code' => $currencyInfo['currency_code'],
                    'currency_symbol' => $currencyInfo['currency_symbol'],
                    'currency_decimal_places' => $currencyInfo['currency_decimal_places'],
                    'in' => $incomes['sums'][$currencyId]['sum'] ?? '0',
                    'out' => $expenses['sums'][$currencyId]['sum'] ?? '0',
                    'sum' => '0',
                ];
                $sums[$currencyId]['sum'] = bcadd($sums[$currencyId]['in'], $sums[$currencyId]['out']);
            }
            
            return ['operations' => $sums];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect balance data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect chart data
     */
    private function collectChartData(): array
    {
        try {
            $controller = app(ChartReportController::class);
            
            $charts = [];
            
            // Account balance chart - try common chart method names
            if (method_exists($controller, 'operations')) {
                $response = $controller->operations($this->accounts, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['operations'] = $response->getData(true);
                }
            }

            // Net worth chart  
            if (method_exists($controller, 'netWorth')) {
                $response = $controller->netWorth($this->accounts, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['net_worth'] = $response->getData(true);
                }
            }

            return $charts;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect chart data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect budget data
     */
    private function collectBudgetData(): array
    {
        if ($this->budgets->count() === 0) {
            return [];
        }

        try {
            $controller = app(BudgetController::class);
            $response = $controller->budgets($this->accounts, $this->budgets, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            if (is_string($response)) {
                return ['html' => $response];
            }
            
            if (method_exists($response, 'render')) {
                return ['html' => $response->render()];
            }
            
            return ['html' => (string)$response];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect budget data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect budget chart data
     */
    private function collectBudgetChartData(): array
    {
        if ($this->budgets->count() === 0) {
            return [];
        }

        try {
            $controller = app(BudgetReportController::class);
            
            $charts = [];
            
            // Budget spending chart
            if (method_exists($controller, 'budgetSpending')) {
                $response = $controller->budgetSpending($this->accounts, $this->budgets, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['budget_spending'] = $response->getData(true);
                }
            }

            return $charts;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect budget chart data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect category data
     */
    private function collectCategoryData(): array
    {
        if ($this->categories->count() === 0) {
            return [];
        }

        try {
            $controller = app(CategoryController::class);
            $response = $controller->categories($this->accounts, $this->categories, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            return ['html' => $response->render()];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect category data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect category chart data
     */
    private function collectCategoryChartData(): array
    {
        if ($this->categories->count() === 0) {
            return [];
        }

        try {
            $controller = app(CategoryReportController::class);
            
            $charts = [];
            
            // Category spending chart
            if (method_exists($controller, 'categorySpending')) {
                $response = $controller->categorySpending($this->accounts, $this->categories, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['category_spending'] = $response->getData(true);
                }
            }

            return $charts;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect category chart data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect tag data
     */
    private function collectTagData(): array
    {
        if ($this->tags->count() === 0) {
            return [];
        }

        try {
            $controller = app(TagController::class);
            $response = $controller->tags($this->accounts, $this->tags, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            return ['html' => $response->render()];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect tag data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect tag chart data
     */
    private function collectTagChartData(): array
    {
        if ($this->tags->count() === 0) {
            return [];
        }

        try {
            $controller = app(TagReportController::class);
            
            $charts = [];
            
            // Tag spending chart
            if (method_exists($controller, 'tagSpending')) {
                $response = $controller->tagSpending($this->accounts, $this->tags, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['tag_spending'] = $response->getData(true);
                }
            }

            return $charts;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect tag chart data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect double report data
     */
    private function collectDoubleData(): array
    {
        if ($this->expenseAccounts->count() === 0) {
            return [];
        }

        try {
            $controller = app(DoubleController::class);
            $response = $controller->accounts($this->accounts, $this->expenseAccounts, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            return ['html' => $response->render()];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect double data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect double chart data
     */
    private function collectDoubleChartData(): array
    {
        if ($this->expenseAccounts->count() === 0) {
            return [];
        }

        try {
            $controller = app(DoubleReportController::class);
            
            $charts = [];
            
            // Double report chart
            if (method_exists($controller, 'doubleReport')) {
                $response = $controller->doubleReport($this->accounts, $this->expenseAccounts, $this->start, $this->end);
                if ($response instanceof \Illuminate\Http\JsonResponse) {
                    $charts['double_report'] = $response->getData(true);
                }
            }

            return $charts;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect double chart data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect income data
     */
    private function collectIncomeData(): array
    {
        try {
            // Get raw data directly from AccountTasker
            $tasker = app(AccountTaskerInterface::class);
            $tasker->setUser($this->user);
            
            return $tasker->getIncomeReport($this->start, $this->end, $this->accounts);
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect income data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect expenses data
     */
    private function collectExpensesData(): array
    {
        try {
            // Get raw data directly from AccountTasker
            $tasker = app(AccountTaskerInterface::class);
            $tasker->setUser($this->user);
            
            return $tasker->getExpenseReport($this->start, $this->end, $this->accounts);
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect expenses data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect operations data
     */
    private function collectOperationsData(): array
    {
        try {
            $data = [];
            
            // Get operations data from balance collection (which now has raw data)
            $balanceData = $this->collectBalanceData();
            if (isset($balanceData['operations'])) {
                $data['operations'] = $balanceData['operations'];
            }

            return $data;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect operations data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    /**
     * Collect bill data
     */
    private function collectBillData(): array
    {
        try {
            $controller = app(BillController::class);
            $response = $controller->bills($this->accounts, $this->start, $this->end);
            
            if ($response instanceof \Illuminate\Http\JsonResponse) {
                return $response->getData(true);
            }
            
            if (is_string($response)) {
                return ['html' => $response];
            }
            
            if (method_exists($response, 'render')) {
                return ['html' => $response->render()];
            }
            
            return ['html' => (string)$response];
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect bill data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }

    // Setters
    public function setUser(User $user): void
    {
        $this->user = $user;
        
        // Set user context for all controllers
        auth()->login($user);
    }

    public function setReportType(string $reportType): void
    {
        $this->reportType = $reportType;
    }

    public function setDateRange(Carbon $start, Carbon $end): void
    {
        $this->start = $start;
        $this->end = $end;
    }

    public function setAccounts(Collection $accounts): void
    {
        $this->accounts = $accounts;
    }

    public function setBudgets(Collection $budgets): void
    {
        $this->budgets = $budgets;
    }

    public function setCategories(Collection $categories): void
    {
        $this->categories = $categories;
    }

    public function setTags(Collection $tags): void
    {
        $this->tags = $tags;
    }

    public function setExpenseAccounts(Collection $expenseAccounts): void
    {
        $this->expenseAccounts = $expenseAccounts;
    }

    /**
     * Collect category operations data for default reports
     */
    private function collectCategoryOperationsData(): array
    {
        try {
            $generator = app(CategoryReportGenerator::class);
            $generator->setUser($this->user);
            $generator->setStart($this->start);
            $generator->setEnd($this->end);
            $generator->setAccounts($this->accounts);
            
            // Generate the report data like CategoryController does
            $generator->operations();
            return $generator->getReport();
        } catch (\Exception $e) {
            app('log')->error(sprintf('Could not collect category operations data: %s', $e->getMessage()));
            return ['error' => $e->getMessage()];
        }
    }
}