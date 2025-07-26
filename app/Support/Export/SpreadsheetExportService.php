<?php

/**
 * SpreadsheetExportService.php
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
use FireflyIII\Exceptions\FireflyException;
use FireflyIII\User;
use Illuminate\Support\Collection;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Chart\Layout;

/**
 * Class SpreadsheetExportService
 * 
 * Handles Excel spreadsheet generation for Firefly III reports
 */
class SpreadsheetExportService
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
    private array $reportData;
    private Spreadsheet $spreadsheet;
    
    // Firefly III brand colors
    private const string PRIMARY_COLOR = '357CA4';
    private const string SUCCESS_COLOR = '28A745';
    private const string DANGER_COLOR = 'DC3545';
    private const string WARNING_COLOR = 'FFC107';
    
    public function __construct()
    {
        $this->accounts = new Collection();
        $this->budgets = new Collection();
        $this->categories = new Collection();
        $this->tags = new Collection();
        $this->expenseAccounts = new Collection();
        $this->reportData = [];
    }

    /**
     * Generate the complete Excel spreadsheet
     * 
     * @throws FireflyException
     */
    public function generateSpreadsheet(): string
    {
        try {
            $this->spreadsheet = new Spreadsheet();
            $this->setupSpreadsheetProperties();
            
            // Create summary sheet
            $this->createSummarySheet();
            
            // Create data sheets based on report type
            $this->createDataSheets();
            
            // Create chart sheets based on report type and available data
            $this->createChartSheets();
            
            // Generate filename and save
            $filename = $this->generateFilename();
            $this->saveSpreadsheet($filename);
            
            return $filename;
            
        } catch (\Exception $e) {
            throw new FireflyException(sprintf('Could not generate report export: %s', $e->getMessage()), 0, $e);
        }
    }

    /**
     * Set up basic spreadsheet properties
     */
    private function setupSpreadsheetProperties(): void
    {
        $this->spreadsheet->getProperties()
            ->setCreator('Firefly III')
            ->setLastModifiedBy('Firefly III')
            ->setTitle(sprintf('Firefly III %s Report', ucfirst($this->reportType)))
            ->setSubject(sprintf('Financial Report: %s to %s', 
                $this->start->format('Y-m-d'), 
                $this->end->format('Y-m-d')
            ))
            ->setDescription(sprintf('Generated on %s by Firefly III', now()->toDateTimeString()))
            ->setKeywords('firefly-iii financial report export')
            ->setCategory('Financial Report');
    }

    /**
     * Create the summary sheet with key metrics
     */
    private function createSummarySheet(): void
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle('Summary');
        
        // Header
        $sheet->setCellValue('A1', 'Firefly III Report Export');
        $sheet->setCellValue('A2', sprintf('%s Report', ucfirst($this->reportType)));
        $sheet->setCellValue('A3', sprintf('Period: %s to %s', 
            $this->start->format('Y-m-d'), 
            $this->end->format('Y-m-d')
        ));
        $sheet->setCellValue('A4', sprintf('Generated: %s', now()->toDateTimeString()));
        $sheet->setCellValue('A5', sprintf('User: %s', $this->user->email));
        
        // Style the header
        $this->styleHeader($sheet, 'A1:E5');
        
        // Set column widths
        $sheet->getColumnDimension('A')->setWidth(20);
        $sheet->getColumnDimension('B')->setWidth(15);
        $sheet->getColumnDimension('C')->setWidth(15);
        $sheet->getColumnDimension('D')->setWidth(15);
        $sheet->getColumnDimension('E')->setWidth(15);
    }

    /**
     * Create data sheets based on report type
     */
    private function createDataSheets(): void
    {
        switch ($this->reportType) {
            case 'budget':
                $this->createBudgetDataSheets();
                break;
            case 'category':
                $this->createCategoryDataSheets();
                break;
            case 'tag':
                $this->createTagDataSheets();
                break;
            case 'double':
                $this->createDoubleDataSheets();
                break;
            case 'audit':
                $this->createAuditDataSheets();
                break;
            default:
                $this->createDefaultDataSheets();
                break;
        }
    }

    /**
     * Create data sheets for default report type
     */
    private function createDefaultDataSheets(): void
    {
        // Account Balances sheet
        $accountSheet = $this->spreadsheet->createSheet();
        $accountSheet->setTitle('Account Balances');
        
        // Set up headers
        $accountSheet->setCellValue('A1', 'Account');
        $accountSheet->setCellValue('B1', 'Start Balance');
        $accountSheet->setCellValue('C1', 'End Balance');
        $accountSheet->setCellValue('D1', 'Difference');
        $accountSheet->setCellValue('E1', 'Currency');
        
        $this->styleHeader($accountSheet, 'A1:E1');
        
        // Populate account data if available
        if (isset($this->reportData['accounts'])) {
            $this->populateAccountData($accountSheet, $this->reportData['accounts']);
        }
        
        // Income/Expense sheet
        $incomeExpenseSheet = $this->spreadsheet->createSheet();
        $incomeExpenseSheet->setTitle('Income vs Expenses');
        
        $incomeExpenseSheet->setCellValue('A1', 'Type');
        $incomeExpenseSheet->setCellValue('B1', 'Amount');
        $incomeExpenseSheet->setCellValue('C1', 'Currency');
        
        $this->styleHeader($incomeExpenseSheet, 'A1:C1');
        
        // Populate balance data if available
        if (isset($this->reportData['balance'])) {
            $this->populateBalanceData($incomeExpenseSheet, $this->reportData['balance']);
        }
        
        // Set column widths
        foreach ([$accountSheet, $incomeExpenseSheet] as $sheet) {
            $sheet->getColumnDimension('A')->setWidth(25);
            $sheet->getColumnDimension('B')->setWidth(15);
            $sheet->getColumnDimension('C')->setWidth(15);
            $sheet->getColumnDimension('D')->setWidth(15);
            $sheet->getColumnDimension('E')->setWidth(10);
        }
    }

    /**
     * Create data sheets for budget report type
     */
    private function createBudgetDataSheets(): void
    {
        // Budget Performance sheet
        $budgetSheet = $this->spreadsheet->createSheet();
        $budgetSheet->setTitle('Budget Performance');
        
        $budgetSheet->setCellValue('A1', 'Budget');
        $budgetSheet->setCellValue('B1', 'Budgeted');
        $budgetSheet->setCellValue('C1', 'Spent');
        $budgetSheet->setCellValue('D1', 'Difference');
        $budgetSheet->setCellValue('E1', 'Currency');
        
        $this->styleHeader($budgetSheet, 'A1:E1');
    }

    /**
     * Create data sheets for category report type
     */
    private function createCategoryDataSheets(): void
    {
        // Category Analysis sheet
        $categorySheet = $this->spreadsheet->createSheet();
        $categorySheet->setTitle('Category Analysis');
        
        $categorySheet->setCellValue('A1', 'Category');
        $categorySheet->setCellValue('B1', 'Income');
        $categorySheet->setCellValue('C1', 'Expenses');
        $categorySheet->setCellValue('D1', 'Difference');
        $categorySheet->setCellValue('E1', 'Currency');
        
        $this->styleHeader($categorySheet, 'A1:E1');
    }

    /**
     * Create data sheets for tag report type
     */
    private function createTagDataSheets(): void
    {
        // Tag Analysis sheet
        $tagSheet = $this->spreadsheet->createSheet();
        $tagSheet->setTitle('Tag Analysis');
        
        $tagSheet->setCellValue('A1', 'Tag');
        $tagSheet->setCellValue('B1', 'Income');
        $tagSheet->setCellValue('C1', 'Expenses');
        $tagSheet->setCellValue('D1', 'Difference');
        $tagSheet->setCellValue('E1', 'Currency');
        
        $this->styleHeader($tagSheet, 'A1:E1');
    }

    /**
     * Create data sheets for double report type
     */
    private function createDoubleDataSheets(): void
    {
        // Asset vs Expense Comparison sheet
        $doubleSheet = $this->spreadsheet->createSheet();
        $doubleSheet->setTitle('Asset vs Expense');
        
        $doubleSheet->setCellValue('A1', 'Asset Account');
        $doubleSheet->setCellValue('B1', 'Expense Account');
        $doubleSheet->setCellValue('C1', 'Amount');
        $doubleSheet->setCellValue('D1', 'Currency');
        
        $this->styleHeader($doubleSheet, 'A1:D1');
    }

    /**
     * Create data sheets for audit report type
     */
    private function createAuditDataSheets(): void
    {
        // Transaction Audit sheet
        $auditSheet = $this->spreadsheet->createSheet();
        $auditSheet->setTitle('Transaction Audit');
        
        $auditSheet->setCellValue('A1', 'Date');
        $auditSheet->setCellValue('B1', 'Description');
        $auditSheet->setCellValue('C1', 'Source');
        $auditSheet->setCellValue('D1', 'Destination');
        $auditSheet->setCellValue('E1', 'Amount');
        $auditSheet->setCellValue('F1', 'Currency');
        
        $this->styleHeader($auditSheet, 'A1:F1');
    }

    /**
     * Apply header styling to a range
     */
    private function styleHeader($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'],
                'size' => 12
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => self::PRIMARY_COLOR]
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '000000']
                ]
            ]
        ]);
    }

    /**
     * Generate filename for the export
     */
    private function generateFilename(): string
    {
        $timestamp = now()->format('Y-m-d_H-i-s');
        $filename = sprintf(
            'FireflyIII_%sReport_%s_to_%s_%s.xlsx',
            ucfirst($this->reportType),
            $this->start->format('Y-m-d'),
            $this->end->format('Y-m-d'),
            $timestamp
        );
        
        return storage_path('app/export/' . $filename);
    }

    /**
     * Save the spreadsheet to disk
     */
    private function saveSpreadsheet(string $filename): void
    {
        // Ensure export directory exists
        $directory = dirname($filename);
        if (!is_dir($directory)) {
            mkdir($directory, 0755, true);
        }
        
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($filename);
    }

    // Setters
    public function setUser(User $user): void
    {
        $this->user = $user;
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

    public function setReportData(array $reportData): void
    {
        $this->reportData = $reportData;
    }

    /**
     * Populate account data in the spreadsheet
     */
    private function populateAccountData($sheet, array $accountData): void
    {
        $row = 2;
        
        // Handle HTML data by extracting useful information
        if (isset($accountData['html'])) {
            // For now, just add a note that this is HTML data
            $sheet->setCellValue('A' . $row, 'Data available in HTML format');
            $sheet->setCellValue('B' . $row, 'See original report');
            return;
        }
        
        // If we have structured data, use it
        if (isset($accountData['accounts'])) {
            foreach ($accountData['accounts'] as $account) {
                $sheet->setCellValue('A' . $row, $account['name'] ?? 'N/A');
                $sheet->setCellValue('B' . $row, $this->formatCurrency($account['start_balance'] ?? 0));
                $sheet->setCellValue('C' . $row, $this->formatCurrency($account['end_balance'] ?? 0));
                $sheet->setCellValue('D' . $row, $this->formatCurrency(($account['end_balance'] ?? 0) - ($account['start_balance'] ?? 0)));
                $sheet->setCellValue('E' . $row, $account['currency_symbol'] ?? '');
                
                // Apply number formatting for currency columns
                $sheet->getStyle('B' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                $sheet->getStyle('C' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                $sheet->getStyle('D' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                
                // Apply color coding for difference column
                $difference = ($account['end_balance'] ?? 0) - ($account['start_balance'] ?? 0);
                if ($difference > 0) {
                    $sheet->getStyle('D' . $row)->getFont()->getColor()->setARGB(self::SUCCESS_COLOR);
                } elseif ($difference < 0) {
                    $sheet->getStyle('D' . $row)->getFont()->getColor()->setARGB(self::DANGER_COLOR);
                }
                
                $row++;
            }
        }
    }

    /**
     * Populate balance data in the spreadsheet
     */
    private function populateBalanceData($sheet, array $balanceData): void
    {
        $row = 2;
        
        // Handle HTML data by extracting useful information
        if (isset($balanceData['html'])) {
            // For now, just add a note that this is HTML data
            $sheet->setCellValue('A' . $row, 'Data available in HTML format');
            $sheet->setCellValue('B' . $row, 'See original report');
            return;
        }
        
        // If we have structured data, use it
        if (isset($balanceData['income'])) {
            $sheet->setCellValue('A' . $row, 'Income');
            $sheet->setCellValue('B' . $row, $this->formatCurrency($balanceData['income']));
            $sheet->setCellValue('C' . $row, $balanceData['currency_symbol'] ?? '');
            $sheet->getStyle('B' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->getStyle('B' . $row)->getFont()->getColor()->setARGB(self::SUCCESS_COLOR);
            $row++;
        }
        
        if (isset($balanceData['expenses'])) {
            $sheet->setCellValue('A' . $row, 'Expenses');
            $sheet->setCellValue('B' . $row, $this->formatCurrency($balanceData['expenses']));
            $sheet->setCellValue('C' . $row, $balanceData['currency_symbol'] ?? '');
            $sheet->getStyle('B' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->getStyle('B' . $row)->getFont()->getColor()->setARGB(self::DANGER_COLOR);
            $row++;
        }
        
        if (isset($balanceData['difference'])) {
            $sheet->setCellValue('A' . $row, 'Net Result');
            $sheet->setCellValue('B' . $row, $this->formatCurrency($balanceData['difference']));
            $sheet->setCellValue('C' . $row, $balanceData['currency_symbol'] ?? '');
            $sheet->getStyle('B' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            
            // Apply color coding based on positive/negative
            $difference = $balanceData['difference'];
            if ($difference > 0) {
                $sheet->getStyle('B' . $row)->getFont()->getColor()->setARGB(self::SUCCESS_COLOR);
            } elseif ($difference < 0) {
                $sheet->getStyle('B' . $row)->getFont()->getColor()->setARGB(self::DANGER_COLOR);
            }
            
            // Make the net result row bold
            $sheet->getStyle('A' . $row . ':C' . $row)->getFont()->setBold(true);
            $row++;
        }
    }

    /**
     * Format currency values for display
     */
    private function formatCurrency(float $value): float
    {
        return round($value, 2);
    }

    /**
     * Create chart sheets based on available chart data
     */
    private function createChartSheets(): void
    {
        if (!isset($this->reportData['charts']) || empty($this->reportData['charts'])) {
            return;
        }

        $this->createGeneralCharts();
        
        // Create report-type specific charts
        switch ($this->reportType) {
            case 'budget':
                $this->createBudgetCharts();
                break;
            case 'category':
                $this->createCategoryCharts();
                break;
            case 'tag':
                $this->createTagCharts();
                break;
            case 'double':
                $this->createDoubleCharts();
                break;
            case 'audit':
                // Audit reports typically don't have charts
                break;
            default:
                $this->createDefaultCharts();
                break;
        }
    }

    /**
     * Create general charts available for all report types
     */
    private function createGeneralCharts(): void
    {
        $chartData = $this->reportData['charts'] ?? [];
        
        // Create operations chart if available
        if (isset($chartData['operations'])) {
            $this->createOperationsChart($chartData['operations']);
        }
        
        // Create net worth chart if available
        if (isset($chartData['net_worth'])) {
            $this->createNetWorthChart($chartData['net_worth']);
        }
    }

    /**
     * Create operations chart
     */
    private function createOperationsChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Operations Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Operations Over Time');
            
            // Create line chart for operations
            $chart = $this->createLineChart(
                $dataWorksheet,
                $chartData,
                'Operations Over Time',
                'Date',
                'Amount'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create operations chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create net worth chart
     */
    private function createNetWorthChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Net Worth Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Net Worth Over Time');
            
            // Create line chart for net worth
            $chart = $this->createLineChart(
                $dataWorksheet,
                $chartData,
                'Net Worth Over Time',
                'Date',
                'Net Worth'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create net worth chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create budget-specific charts
     */
    private function createBudgetCharts(): void
    {
        if (!isset($this->reportData['budget_charts'])) {
            return;
        }

        $chartData = $this->reportData['budget_charts'];
        
        if (isset($chartData['budget_spending'])) {
            $this->createBudgetSpendingChart($chartData['budget_spending']);
        }
    }

    /**
     * Create budget spending chart
     */
    private function createBudgetSpendingChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Budget Spending Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Budget Spending Analysis');
            
            // Create pie chart for budget spending
            $chart = $this->createPieChart(
                $dataWorksheet,
                $chartData,
                'Budget Spending Distribution'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create budget spending chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create category-specific charts
     */
    private function createCategoryCharts(): void
    {
        if (!isset($this->reportData['category_charts'])) {
            return;
        }

        $chartData = $this->reportData['category_charts'];
        
        if (isset($chartData['category_spending'])) {
            $this->createCategorySpendingChart($chartData['category_spending']);
        }
    }

    /**
     * Create category spending chart
     */
    private function createCategorySpendingChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Category Spending Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Category Spending Analysis');
            
            // Create pie chart for category spending
            $chart = $this->createPieChart(
                $dataWorksheet,
                $chartData,
                'Category Spending Distribution'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create category spending chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create tag-specific charts
     */
    private function createTagCharts(): void
    {
        if (!isset($this->reportData['tag_charts'])) {
            return;
        }

        $chartData = $this->reportData['tag_charts'];
        
        if (isset($chartData['tag_spending'])) {
            $this->createTagSpendingChart($chartData['tag_spending']);
        }
    }

    /**
     * Create tag spending chart
     */
    private function createTagSpendingChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Tag Spending Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Tag Spending Analysis');
            
            // Create pie chart for tag spending
            $chart = $this->createPieChart(
                $dataWorksheet,
                $chartData,
                'Tag Spending Distribution'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create tag spending chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create double report charts
     */
    private function createDoubleCharts(): void
    {
        if (!isset($this->reportData['double_charts'])) {
            return;
        }

        $chartData = $this->reportData['double_charts'];
        
        if (isset($chartData['double_report'])) {
            $this->createDoubleReportChart($chartData['double_report']);
        }
    }

    /**
     * Create double report chart
     */
    private function createDoubleReportChart(array $chartData): void
    {
        if (!$this->isValidChartData($chartData)) {
            return;
        }

        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Asset vs Expense Chart');
            
            // Prepare data for the chart
            $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Asset vs Expense Analysis');
            
            // Create bar chart for double report
            $chart = $this->createBarChart(
                $dataWorksheet,
                $chartData,
                'Asset vs Expense Comparison'
            );
            
            if ($chart !== null) {
                $chartSheet->addChart($chart);
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create double report chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create default charts for standard reports
     */
    private function createDefaultCharts(): void
    {
        // Default charts are handled by createGeneralCharts()
    }

    /**
     * Validate chart data structure
     */
    private function isValidChartData(array $chartData): bool
    {
        return !empty($chartData) && 
               (isset($chartData['datasets']) || isset($chartData['labels']));
    }

    /**
     * Prepare chart data in worksheet
     */
    private function prepareChartData($worksheet, array $chartData, string $title)
    {
        // Add title
        $worksheet->setCellValue('A1', $title);
        $this->styleHeader($worksheet, 'A1:C1');
        
        $row = 3;
        
        // Handle different chart data formats
        if (isset($chartData['labels']) && isset($chartData['datasets'])) {
            // ChartJS format
            $labels = $chartData['labels'];
            $datasets = $chartData['datasets'];
            
            // Headers
            $worksheet->setCellValue('A2', 'Label');
            $col = 'B';
            foreach ($datasets as $dataset) {
                $worksheet->setCellValue($col . '2', $dataset['label'] ?? 'Data');
                $col++;
            }
            
            // Data
            foreach ($labels as $index => $label) {
                $worksheet->setCellValue('A' . $row, $label);
                $col = 'B';
                foreach ($datasets as $dataset) {
                    $value = $dataset['data'][$index] ?? 0;
                    $worksheet->setCellValue($col . $row, $this->formatCurrency((float)$value));
                    $col++;
                }
                $row++;
            }
        } elseif (is_array($chartData) && !empty($chartData)) {
            // Simple key-value format
            $worksheet->setCellValue('A2', 'Category');
            $worksheet->setCellValue('B2', 'Value');
            
            foreach ($chartData as $key => $value) {
                $worksheet->setCellValue('A' . $row, $key);
                $worksheet->setCellValue('B' . $row, $this->formatCurrency((float)$value));
                $row++;
            }
        }
        
        // Style the headers
        $this->styleHeader($worksheet, 'A2:C2');
        
        return $worksheet;
    }

    /**
     * Create a line chart
     */
    private function createLineChart($worksheet, array $chartData, string $title, string $xAxisTitle = '', string $yAxisTitle = ''): ?Chart
    {
        if (!$this->isValidChartData($chartData)) {
            return null;
        }

        try {
            $dataSeriesLabels = [];
            $dataSeriesValues = [];
            $dataRange = '';
            
            if (isset($chartData['datasets']) && !empty($chartData['datasets'])) {
                $datasets = $chartData['datasets'];
                $rowCount = count($chartData['labels'] ?? []);
                
                foreach ($datasets as $index => $dataset) {
                    $colLetter = chr(66 + $index); // B, C, D, etc.
                    
                    // Series label
                    $dataSeriesLabels[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_STRING,
                        $worksheet->getTitle() . '!' . '$' . $colLetter . '$2',
                        null,
                        1
                    );
                    
                    // Series values
                    $dataSeriesValues[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_NUMBER,
                        $worksheet->getTitle() . '!' . '$' . $colLetter . '$3:$' . $colLetter . '$' . (2 + $rowCount),
                        null,
                        $rowCount
                    );
                }
                
                // Category labels (X-axis)
                $xAxisRange = $worksheet->getTitle() . '!' . '$A$3:$A$' . (2 + $rowCount);
            }
            
            // Create data series
            $xAxisLabels = isset($xAxisRange) ? new DataSeriesValues(
                DataSeriesValues::DATASERIES_TYPE_STRING,
                $xAxisRange,
                null,
                count($chartData['labels'] ?? [])
            ) : new DataSeriesValues();
            
            $series = new DataSeries(
                DataSeries::TYPE_LINECHART,
                DataSeries::GROUPING_STANDARD,
                range(0, count($dataSeriesValues) - 1),
                $dataSeriesLabels,
                [$xAxisLabels],
                $dataSeriesValues
            );
            
            // Create plot area
            $plotArea = new PlotArea(null, [$series]);
            
            // Create legend
            $legend = new Legend(Legend::POSITION_RIGHT, null, false);
            
            // Create chart title
            $chartTitle = new Title($title);
            
            // Create chart
            $chart = new Chart(
                'chart_' . uniqid(),
                $chartTitle,
                $legend,
                $plotArea,
                true,
                DataSeries::EMPTY_AS_GAP,
                null,
                null
            );
            
            // Set chart position
            $chart->setTopLeftPosition('A5');
            $chart->setBottomRightPosition('H20');
            
            return $chart;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create line chart: %s', $e->getMessage()));
            return null;
        }
    }

    /**
     * Create a pie chart
     */
    private function createPieChart($worksheet, array $chartData, string $title): ?Chart
    {
        if (!$this->isValidChartData($chartData)) {
            return null;
        }

        try {
            // Determine data range
            $lastRow = 2;
            $col = 'A';
            while ($worksheet->getCell($col . ($lastRow + 1))->getValue() !== null) {
                $lastRow++;
            }
            
            if ($lastRow < 3) {
                return null; // No data
            }
            
            // Create data series for pie chart
            $dataSeriesLabels = [
                new DataSeriesValues(
                    DataSeriesValues::DATASERIES_TYPE_STRING,
                    $worksheet->getTitle() . '!' . '$B$2',
                    null,
                    1
                )
            ];
            
            $dataSeriesValues = [
                new DataSeriesValues(
                    DataSeriesValues::DATASERIES_TYPE_NUMBER,
                    $worksheet->getTitle() . '!' . '$B$3:$B$' . $lastRow,
                    null,
                    $lastRow - 2
                )
            ];
            
            $xAxisLabels = new DataSeriesValues(
                DataSeriesValues::DATASERIES_TYPE_STRING,
                $worksheet->getTitle() . '!' . '$A$3:$A$' . $lastRow,
                null,
                $lastRow - 2
            );
            
            // Create data series
            $series = new DataSeries(
                DataSeries::TYPE_PIECHART,
                null,
                range(0, count($dataSeriesValues) - 1),
                $dataSeriesLabels,
                [$xAxisLabels],
                $dataSeriesValues
            );
            
            // Create plot area
            $plotArea = new PlotArea(null, [$series]);
            
            // Create legend
            $legend = new Legend(Legend::POSITION_RIGHT, null, false);
            
            // Create chart title
            $chartTitle = new Title($title);
            
            // Create chart
            $chart = new Chart(
                'chart_' . uniqid(),
                $chartTitle,
                $legend,
                $plotArea,
                true,
                DataSeries::EMPTY_AS_GAP,
                null,
                null
            );
            
            // Set chart position
            $chart->setTopLeftPosition('A5');
            $chart->setBottomRightPosition('H20');
            
            return $chart;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create pie chart: %s', $e->getMessage()));
            return null;
        }
    }

    /**
     * Create a bar chart
     */
    private function createBarChart($worksheet, array $chartData, string $title): ?Chart
    {
        if (!$this->isValidChartData($chartData)) {
            return null;
        }

        try {
            $dataSeriesLabels = [];
            $dataSeriesValues = [];
            
            if (isset($chartData['datasets']) && !empty($chartData['datasets'])) {
                $datasets = $chartData['datasets'];
                $rowCount = count($chartData['labels'] ?? []);
                
                foreach ($datasets as $index => $dataset) {
                    $colLetter = chr(66 + $index); // B, C, D, etc.
                    
                    // Series label
                    $dataSeriesLabels[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_STRING,
                        $worksheet->getTitle() . '!' . '$' . $colLetter . '$2',
                        null,
                        1
                    );
                    
                    // Series values
                    $dataSeriesValues[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_NUMBER,
                        $worksheet->getTitle() . '!' . '$' . $colLetter . '$3:$' . $colLetter . '$' . (2 + $rowCount),
                        null,
                        $rowCount
                    );
                }
                
                // Category labels (X-axis)
                $xAxisRange = $worksheet->getTitle() . '!' . '$A$3:$A$' . (2 + $rowCount);
            }
            
            // Create data series
            $xAxisLabels = isset($xAxisRange) ? new DataSeriesValues(
                DataSeriesValues::DATASERIES_TYPE_STRING,
                $xAxisRange,
                null,
                count($chartData['labels'] ?? [])
            ) : new DataSeriesValues();
            
            $series = new DataSeries(
                DataSeries::TYPE_BARCHART,
                DataSeries::GROUPING_CLUSTERED,
                range(0, count($dataSeriesValues) - 1),
                $dataSeriesLabels,
                [$xAxisLabels],
                $dataSeriesValues
            );
            
            // Create plot area
            $plotArea = new PlotArea(null, [$series]);
            
            // Create legend
            $legend = new Legend(Legend::POSITION_RIGHT, null, false);
            
            // Create chart title
            $chartTitle = new Title($title);
            
            // Create chart
            $chart = new Chart(
                'chart_' . uniqid(),
                $chartTitle,
                $legend,
                $plotArea,
                true,
                DataSeries::EMPTY_AS_GAP,
                null,
                null
            );
            
            // Set chart position
            $chart->setTopLeftPosition('A5');
            $chart->setBottomRightPosition('H20');
            
            return $chart;
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create bar chart: %s', $e->getMessage()));
            return null;
        }
    }
}