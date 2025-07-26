<?php

/**
 * BalanceController.php
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

namespace FireflyIII\Http\Controllers\Report;

use Throwable;
use Carbon\Carbon;
use FireflyIII\Enums\TransactionTypeEnum;
use FireflyIII\Exceptions\FireflyException;
use FireflyIII\Helpers\Collector\GroupCollectorInterface;
use FireflyIII\Http\Controllers\Controller;
use FireflyIII\Models\Account;
use FireflyIII\Models\Budget;
use FireflyIII\Repositories\Budget\BudgetRepositoryInterface;
use FireflyIII\Support\CacheProperties;
use Illuminate\Support\Collection;

/**
 * Class BalanceController.
 */
class BalanceController extends Controller
{
    /** @var BudgetRepositoryInterface */
    private $repository;

    /**
     * BalanceController constructor.
     */
    public function __construct()
    {
        parent::__construct();

        $this->middleware(
            function ($request, $next) {
                $this->repository = app(BudgetRepositoryInterface::class);

                return $next($request);
            }
        );
    }

    /**
     * Show overview of budget balances.
     *
     * @return string
     *
     * @throws FireflyException
     */
    public function general(Collection $accounts, Carbon $start, Carbon $end)
    {
        $report  = [
            'budgets'  => [],
            'accounts' => [],
        ];

        /** @var Account $account */
        foreach ($accounts as $account) {
            $report['accounts'][$account->id] = [
                'id'   => $account->id,
                'name' => $account->name,
                'iban' => $account->iban,
                'sum'  => '0',
            ];
        }

        $budgets = $this->repository->getBudgets();

        /** @var Budget $budget */
        foreach ($budgets as $budget) {
            $budgetId                              = $budget->id;
            $report['budgets'][$budgetId]          = [
                'budget_id'   => $budgetId,
                'budget_name' => $budget->name,
                'spent'       => [], // per account
                'sums'        => [], // per currency
            ];
            $spent                                 = [];

            /** @var GroupCollectorInterface $collector */
            $collector                             = app(GroupCollectorInterface::class);
            $journals                              = $collector->setRange($start, $end)->setSourceAccounts($accounts)->setTypes([TransactionTypeEnum::WITHDRAWAL->value])->setBudget($budget)
                ->getExtractedJournals()
            ;

            /** @var array $journal */
            foreach ($journals as $journal) {
                $sourceAccount                                                 = $journal['source_account_id'];
                $currencyId                                                    = $journal['currency_id'];
                $spent[$sourceAccount]                  ??= [
                    'source_account_id'       => $sourceAccount,
                    'currency_id'             => $journal['currency_id'],
                    'currency_code'           => $journal['currency_code'],
                    'currency_name'           => $journal['currency_name'],
                    'currency_symbol'         => $journal['currency_symbol'],
                    'currency_decimal_places' => $journal['currency_decimal_places'],
                    'spent'                   => '0',
                ];
                $spent[$sourceAccount]['spent']                                = bcadd($spent[$sourceAccount]['spent'], (string) $journal['amount']);

                // also fix sum:
                $report['sums'][$budgetId][$currencyId] ??= [
                    'sum'                     => '0',
                    'currency_id'             => $journal['currency_id'],
                    'currency_code'           => $journal['currency_code'],
                    'currency_name'           => $journal['currency_name'],
                    'currency_symbol'         => $journal['currency_symbol'],
                    'currency_decimal_places' => $journal['currency_decimal_places'],
                ];
                $report['sums'][$budgetId][$currencyId]['sum']                 = bcadd($report['sums'][$budgetId][$currencyId]['sum'], (string) $journal['amount']);
                $report['accounts'][$sourceAccount]['sum']                     = bcadd($report['accounts'][$sourceAccount]['sum'], (string) $journal['amount']);

                // add currency info for account sum
                $report['accounts'][$sourceAccount]['currency_id']             = $journal['currency_id'];
                $report['accounts'][$sourceAccount]['currency_code']           = $journal['currency_code'];
                $report['accounts'][$sourceAccount]['currency_name']           = $journal['currency_name'];
                $report['accounts'][$sourceAccount]['currency_symbol']         = $journal['currency_symbol'];
                $report['accounts'][$sourceAccount]['currency_decimal_places'] = $journal['currency_decimal_places'];
            }
            $report['budgets'][$budgetId]['spent'] = $spent;
            // get transactions in budget
        }

        try {
            $result = view('reports.partials.balance', compact('report'))->render();
        } catch (Throwable $e) {
            app('log')->error(sprintf('Could not render reports.partials.balance: %s', $e->getMessage()));
            app('log')->error($e->getTraceAsString());
            $result = 'Could not render view.';

            throw new FireflyException($result, 0, $e);
        }

        return $result;
    }

    /**
     * Return income vs expenses data for Excel export.
     * 
     * @throws FireflyException
     */
    public function incomeVsExpenses(Collection $accounts, Carbon $start, Carbon $end)
    {
        // Chart properties for cache:
        $cache = new CacheProperties();
        $cache->addProperty($start);
        $cache->addProperty($end);
        $cache->addProperty('income-expenses-export');
        $cache->addProperty($accounts->pluck('id')->toArray());
        if ($cache->has()) {
            return response()->json($cache->get());
        }

        // Collect income transactions (deposits)
        /** @var GroupCollectorInterface $incomeCollector */
        $incomeCollector = app(GroupCollectorInterface::class);
        $incomeJournals = $incomeCollector
            ->setRange($start, $end)
            ->setAccounts($accounts)
            ->setTypes([TransactionTypeEnum::DEPOSIT->value])
            ->getExtractedJournals();

        // Collect expense transactions (withdrawals)
        /** @var GroupCollectorInterface $expenseCollector */
        $expenseCollector = app(GroupCollectorInterface::class);
        $expenseJournals = $expenseCollector
            ->setRange($start, $end)
            ->setAccounts($accounts)
            ->setTypes([TransactionTypeEnum::WITHDRAWAL->value])
            ->getExtractedJournals();

        // Group by currency
        $incomeByCurrency = $this->groupByCurrency($incomeJournals);
        $expensesByCurrency = $this->groupByCurrency($expenseJournals);

        // Combine results - for now, focus on the default currency
        $totalIncome = 0;
        $totalExpenses = 0;
        $currencySymbol = '';

        foreach ($incomeByCurrency as $currencyId => $data) {
            $totalIncome += $data['sum'];
            $currencySymbol = $data['currency_symbol'];
        }

        foreach ($expensesByCurrency as $currencyId => $data) {
            $totalExpenses += abs($data['sum']); // expenses are usually negative, make positive
        }

        $exportData = [
            'income' => $totalIncome,
            'expenses' => $totalExpenses,
            'difference' => $totalIncome - $totalExpenses,
            'currency_symbol' => $currencySymbol
        ];

        $cache->store($exportData);

        return response()->json($exportData);
    }

    /**
     * Group journal data by currency
     */
    private function groupByCurrency(array $journals): array
    {
        $grouped = [];

        foreach ($journals as $journal) {
            $currencyId = $journal['currency_id'];

            if (!isset($grouped[$currencyId])) {
                $grouped[$currencyId] = [
                    'currency_id' => $journal['currency_id'],
                    'currency_code' => $journal['currency_code'],
                    'currency_name' => $journal['currency_name'],
                    'currency_symbol' => $journal['currency_symbol'],
                    'currency_decimal_places' => $journal['currency_decimal_places'],
                    'sum' => 0
                ];
            }

            $grouped[$currencyId]['sum'] += (float)$journal['amount'];
        }

        return $grouped;
    }
}
