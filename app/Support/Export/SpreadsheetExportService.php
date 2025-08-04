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
use FireflyIII\Repositories\Account\AccountTaskerInterface;
use FireflyIII\Support\Report\Category\CategoryReportGenerator;

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
    
    // Subtle Google Sheets-inspired blue color scheme for professional appearance
    private const string PRIMARY_COLOR = '9FC5E8';      // Much softer light blue for headers
    private const string HEADER_COLOR = '6FA8DC';        // Softer medium blue for section headers  
    private const string SUCCESS_COLOR = '228B22';       // Forest green for positive values (used consistently)
    private const string DANGER_COLOR = 'F56C6C';        // Softer red for negative values  
    private const string SUCCESS_BRIGHT = '228B22';      // Same as SUCCESS_COLOR for consistency
    private const string WARNING_COLOR = 'F9AB00';       // Orange for expense highlights
    private const string NEUTRAL_COLOR = '5F6368';       // Google gray for neutral values
    private const string LIGHT_BLUE = 'E8F0FE';          // Very light blue for alternating rows (Google style)
    private const string MEDIUM_BLUE = 'D2E3FC';         // Light blue for borders and accents
    private const string DARK_GRAY = '202124';           // Google dark gray for text
    private const string WHITE = 'FFFFFF';               // White for backgrounds
    private const string TOTAL_ROW_COLOR = 'C5E1F5';     // Light blue for totals (matches theme)
    
    // Firefly III Chart Colors - 17-color palette
    private const array CHART_COLORS = [
        '357CA4', // Primary blue (with transparency applied in charts)
        '008D4C', // Green
        'DB8B0B', // Orange/Gold
        'CA195A', // Magenta
        '555299', // Purple
        '4285F4', // Google Blue
        'DB4437', // Red
        'F4B400', // Yellow
        '0F9D58', // Forest Green
        'AB47BC', // Violet
        '00ACC1', // Cyan
        'FF7043', // Orange
        '9E9D24', // Olive
        '5C6BC0', // Periwinkle
        'F06292', // Pink
        '00796B', // Teal
        'C2185B', // Deep Pink
    ];
    
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
     * Create the summary sheet with enhanced branding and key metrics
     */
    private function createSummarySheet(): void
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle('Summary');
        
        // Create professional branding section
        $nextRow = $this->createBrandingSection($sheet, 1);
        
        // Add summary metrics if available
        if (!empty($this->reportData)) {
            $this->addSummaryMetrics($sheet, $nextRow);
        }
        
        // Set dynamic column widths for better appearance
        $this->setDynamicColumnWidths($sheet, [
            'A' => 25,  // Increased for better label display
            'B' => 18,  // Increased for data
            'C' => 15,
            'D' => 18,  // Increased for labels
            'E' => 18   // Increased for data
        ]);
    }
    
    /**
     * Create professional title/branding section
     */
    private function createBrandingSection($sheet, int $startRow = 1): int
    {
        $row = $startRow;
        
        // Main title
        $sheet->setCellValue('A' . $row, 'Firefly III Financial Report');
        $this->styleSectionHeader($sheet, 'A' . $row . ':E' . $row);
        $row++;
        
        // Report type
        $sheet->setCellValue('A' . $row, sprintf('%s Report', ucfirst($this->reportType)));
        $this->styleHeader($sheet, 'A' . $row . ':E' . $row);
        $row++;
        
        // Report details in a structured format
        $sheet->setCellValue('A' . $row, 'Period:');
        $sheet->setCellValue('B' . $row, sprintf('%s to %s', 
            $this->start->format('M j, Y'), 
            $this->end->format('M j, Y')
        ));
        
        $sheet->setCellValue('D' . $row, 'Generated:');
        $sheet->setCellValue('E' . $row, now()->format('M j, Y g:i A'));
        
        $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, true);
        $row++;
        
        // User info
        $sheet->setCellValue('A' . $row, 'User:');
        $sheet->setCellValue('B' . $row, $this->user->email);
        $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, false);
        $row++;
        
        // Add spacing
        $row++;
        
        return $row;
    }
    
    /**
     * Add summary metrics to the summary sheet
     */
    private function addSummaryMetrics($sheet, int $startRow): void
    {
        $row = $startRow;
        
        // Summary section header
        $sheet->setCellValue('A' . $row, 'Report Summary');
        $this->styleHeader($sheet, 'A' . $row . ':E' . $row);
        $row++;
        
        // Add key metrics based on available data
        if (isset($this->reportData['balance'])) {
            $balanceData = $this->reportData['balance'];
            
            if (isset($balanceData['operations'])) {
                foreach ($balanceData['operations'] as $currencyId => $data) {
                    $currency = $data['currency_symbol'] ?? $data['currency_code'] ?? '';
                    
                    // Income row
                    $sheet->setCellValue('A' . $row, 'Total Income (' . $currency . ')');
                    $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($data['in'] ?? 0)));
                    $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, true);
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($data['in'] ?? 0), true);
                    $row++;
                    
                    // Expenses row
                    $sheet->setCellValue('A' . $row, 'Total Expenses (' . $currency . ')');
                    $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($data['out'] ?? 0)));
                    $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, false);
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($data['out'] ?? 0), false);
                    $row++;
                    
                    // Total row using SUM formula
                    $incomeRow = $row - 2; // Income row is 2 rows above
                    $expenseRow = $row - 1; // Expense row is 1 row above
                    $netResult = (float)($data['sum'] ?? 0);
                    $sheet->setCellValue('A' . $row, 'TOTAL (' . $currency . ')');
                    $sheet->setCellValue('B' . $row, sprintf('=B%d+B%d', $incomeRow, $expenseRow));
                    $this->styleTotalRow($sheet, 'A' . $row . ':E' . $row);
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, $netResult, false);
                    $row++;
                    
                    $row++; // Add spacing between currencies
                }
            }
        }
        
        // Add account summary if available
        if (isset($this->reportData['accounts']['accounts']) && count($this->reportData['accounts']['accounts']) > 0) {
            $sheet->setCellValue('A' . $row, 'Number of Accounts');
            $sheet->setCellValue('B' . $row, count($this->reportData['accounts']['accounts']));
            $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, true);
            $row++;
        }
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
        
        // Income sheet
        $incomeSheet = $this->spreadsheet->createSheet();
        $incomeSheet->setTitle('Income');
        
        $incomeSheet->setCellValue('A1', 'Account');
        $incomeSheet->setCellValue('B1', 'Amount');
        $incomeSheet->setCellValue('C1', 'Currency');
        
        $this->styleHeader($incomeSheet, 'A1:C1');
        
        // Populate income data if available
        if (isset($this->reportData['income'])) {
            $this->populateIncomeData($incomeSheet, $this->reportData['income']);
        }
        
        // Expenses sheet
        $expenseSheet = $this->spreadsheet->createSheet();
        $expenseSheet->setTitle('Expenses');
        
        $expenseSheet->setCellValue('A1', 'Account');
        $expenseSheet->setCellValue('B1', 'Amount');
        $expenseSheet->setCellValue('C1', 'Currency');
        
        $this->styleHeader($expenseSheet, 'A1:C1');
        
        // Populate expenses data if available
        if (isset($this->reportData['expenses'])) {
            $this->populateExpensesData($expenseSheet, $this->reportData['expenses']);
        }
        
        // Income vs Expenses Summary sheet
        $summarySheet = $this->spreadsheet->createSheet();
        $summarySheet->setTitle('Income vs Expenses');
        
        $summarySheet->setCellValue('A1', 'Type');
        $summarySheet->setCellValue('B1', 'Amount');
        $summarySheet->setCellValue('C1', 'Currency');
        
        $this->styleHeader($summarySheet, 'A1:C1');
        
        // Populate balance data if available
        if (isset($this->reportData['balance'])) {
            $this->populateBalanceData($summarySheet, $this->reportData['balance']);
        }
        
        // Categories sheet with monthly breakdown
        $categoriesSheet = $this->spreadsheet->createSheet();
        $categoriesSheet->setTitle('Categories');
        
        // Create monthly categories data similar to the web report
        $this->createMonthlyCategoriesSheet($categoriesSheet);
        
        // Set dynamic column widths for better professional appearance
        $this->setDynamicColumnWidths($accountSheet, [
            'A' => 30, 'B' => 18, 'C' => 18, 'D' => 18, 'E' => 12
        ]);
        
        $this->setDynamicColumnWidths($incomeSheet, [
            'A' => 30, 'B' => 18, 'C' => 12
        ]);
        
        $this->setDynamicColumnWidths($expenseSheet, [
            'A' => 30, 'B' => 18, 'C' => 12
        ]);
        
        $this->setDynamicColumnWidths($summarySheet, [
            'A' => 30, 'B' => 18, 'C' => 18
        ]);
        
        // Categories sheet column widths are set in createMonthlyCategoriesSheet method
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
     * Apply professional header styling with gradient effect
     */
    private function styleHeader($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => self::DARK_GRAY],
                'size' => 12,
                'name' => 'Calibri'
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
                'outline' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => self::DARK_GRAY]
                ],
                'inside' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => self::MEDIUM_BLUE]
                ]
            ]
        ]);
        
        // Set row height for better appearance
        $this->setRowHeight($sheet, $range, 22);
    }
    
    /**
     * Apply section header styling (larger, more prominent) - Google Sheets style
     */
    private function styleSectionHeader($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => self::WHITE],
                'size' => 14,
                'name' => 'Calibri'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => self::HEADER_COLOR]
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER
            ],
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['argb' => self::DARK_GRAY]
                ]
            ]
        ]);
        
        $this->setRowHeight($sheet, $range, 26);
    }
    
    /**
     * Apply alternating row styling (Google Sheets-inspired blue theme)
     */
    private function styleAlternatingRows($sheet, string $range, bool $isOddRow = false): void
    {
        $backgroundColor = $isOddRow ? self::WHITE : self::LIGHT_BLUE;
        
        $sheet->getStyle($range)->applyFromArray([
            'font' => [
                'size' => 11,
                'name' => 'Calibri'
                // Removed color override to preserve currency formatting
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => $backgroundColor]
            ],
            'alignment' => [
                'vertical' => Alignment::VERTICAL_CENTER
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_HAIR,
                    'color' => ['argb' => self::MEDIUM_BLUE]
                ]
            ]
        ]);
        
        $this->setRowHeight($sheet, $range, 18);
    }
    
    /**
     * Set default text color for non-currency cells
     */
    private function setDefaultTextColor($sheet, string $range): void
    {
        $sheet->getStyle($range)->getFont()->getColor()->setARGB(self::DARK_GRAY);
    }
    
    /**
     * Apply professional total row styling - Google Sheets inspired
     */
    private function styleTotalRow($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 12,
                'name' => 'Calibri',
                'color' => ['argb' => self::DARK_GRAY]
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => self::TOTAL_ROW_COLOR]
            ],
            'alignment' => [
                'vertical' => Alignment::VERTICAL_CENTER
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => self::DARK_GRAY]
                ],
                'bottom' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => self::DARK_GRAY]
                ],
                'left' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => self::MEDIUM_BLUE]
                ],
                'right' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => self::MEDIUM_BLUE]
                ]
            ]
        ]);
        
        $this->setRowHeight($sheet, $range, 20);
    }
    
    /**
     * Apply currency formatting with conditional colors
     */
    private function applyCurrencyFormatting($sheet, string $range, float $value = null, bool $isIncome = false, bool $isExpense = false): void
    {
        // Apply number formatting
        $sheet->getStyle($range)->getNumberFormat()->setFormatCode('#,##0.00');
        
        // Apply conditional color coding if value is provided
        if ($value !== null) {
            $color = self::NEUTRAL_COLOR; // Default neutral gray
            
            if ($isExpense) {
                // Expenses should always be red (regardless of sign)
                $color = self::DANGER_COLOR;
            } elseif ($isIncome) {
                // Income should be green if positive, neutral if zero/negative
                $color = $value > 0 ? self::SUCCESS_COLOR : self::NEUTRAL_COLOR;
            } else {
                // General case: green for positive, red for negative
                if ($value > 0) {
                    $color = self::SUCCESS_COLOR;
                } elseif ($value < 0) {
                    $color = self::DANGER_COLOR;
                }
            }
            
            $sheet->getStyle($range)->getFont()->getColor()->setARGB($color);
        }
    }
    
    /**
     * Set row height for a range
     */
    private function setRowHeight($sheet, string $range, float $height): void
    {
        // Extract row numbers from range
        if (preg_match('/([0-9]+):.*?([0-9]+)/', $range, $matches)) {
            $startRow = (int)$matches[1];
            $endRow = (int)$matches[2];
            
            for ($row = $startRow; $row <= $endRow; $row++) {
                $sheet->getRowDimension($row)->setRowHeight($height);
            }
        } elseif (preg_match('/[A-Z]+([0-9]+)/', $range, $matches)) {
            // Single cell range
            $row = (int)$matches[1];
            $sheet->getRowDimension($row)->setRowHeight($height);
        }
    }
    
    /**
     * Set dynamic column widths based on content
     */
    private function setDynamicColumnWidths($sheet, array $columnWidths): void
    {
        foreach ($columnWidths as $column => $width) {
            $sheet->getColumnDimension($column)->setWidth($width);
        }
    }
    
    /**
     * Apply professional table styling to a data range
     */
    private function styleDataTable($sheet, string $headerRange, string $dataRange, int $dataStartRow, int $dataEndRow, array $currencyColumns = []): void
    {
        // Style header
        $this->styleHeader($sheet, $headerRange);
        
        // Style data rows with alternating colors
        for ($row = $dataStartRow; $row <= $dataEndRow; $row++) {
            $isOddRow = ($row - $dataStartRow) % 2 === 0;
            $rowRange = preg_replace('/[0-9]+:[A-Z]+[0-9]+/', $row . ':' . substr($dataRange, strpos($dataRange, ':') + 1, 1) . $row, $dataRange);
            $this->styleAlternatingRows($sheet, $rowRange, $isOddRow);
            
            // Apply currency formatting to specified columns
            foreach ($currencyColumns as $column => $isIncome) {
                $cellRange = $column . $row;
                $cellValue = $sheet->getCell($cellRange)->getCalculatedValue();
                if (is_numeric($cellValue)) {
                    $this->applyCurrencyFormatting($sheet, $cellRange, (float)$cellValue, $isIncome);
                }
            }
        }
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
        $writer->setIncludeCharts(true); // Enable chart writing for LibreOffice compatibility
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
     * Populate account data in the spreadsheet with professional styling
     */
    private function populateAccountData($sheet, array $accountData): void
    {
        $row = 2;
        
        // Debug logging
        app('log')->info('Account data received for export', ['data' => $accountData]);
        
        // Try to extract data from different possible formats
        if (isset($accountData['html']) && is_string($accountData['html'])) {
            // Parse HTML to extract account data
            $this->parseAccountDataFromHtml($sheet, $accountData['html']);
            return;
        }
        
        // If we have structured data, use it
        if (isset($accountData['accounts']) && is_array($accountData['accounts'])) {
            $startDataRow = $row;
            foreach ($accountData['accounts'] as $account) {
                $sheet->setCellValue('A' . $row, $account['name'] ?? 'N/A');
                $sheet->setCellValue('B' . $row, $this->formatCurrency($account['start_balance'] ?? 0));
                $sheet->setCellValue('C' . $row, $this->formatCurrency($account['end_balance'] ?? 0));
                $sheet->setCellValue('D' . $row, $this->formatCurrency(($account['end_balance'] ?? 0) - ($account['start_balance'] ?? 0)));
                $sheet->setCellValue('E' . $row, $account['currency_symbol'] ?? '');
                
                // Apply professional currency formatting with conditional colors
                $startBalance = (float)($account['start_balance'] ?? 0);
                $endBalance = (float)($account['end_balance'] ?? 0);
                $difference = $endBalance - $startBalance;
                
                $this->applyCurrencyFormatting($sheet, 'B' . $row, $startBalance, false);
                $this->applyCurrencyFormatting($sheet, 'C' . $row, $endBalance, false);
                $this->applyCurrencyFormatting($sheet, 'D' . $row, $difference, false);
                
                // Apply alternating row styling
                $isOddRow = ($row - $startDataRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, $isOddRow);
                
                // Set default text color for non-currency columns (name and currency symbol)
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'E' . $row);
                
                $row++;
            }
            
            // Add totals grouped by currency if we have multiple accounts
            if ($row > $startDataRow + 1) {
                $row++; // Empty separator row
                $endDataRow = $row - 2;
                
                // Group accounts by currency for separate totals
                $currencyTotals = [];
                foreach ($accountData['accounts'] as $account) {
                    $currency = $account['currency_symbol'] ?? 'Unknown';
                    if (!isset($currencyTotals[$currency])) {
                        $currencyTotals[$currency] = [
                            'start_balance' => 0,
                            'end_balance' => 0,
                            'difference' => 0
                        ];
                    }
                    $currencyTotals[$currency]['start_balance'] += (float)($account['start_balance'] ?? 0);
                    $currencyTotals[$currency]['end_balance'] += (float)($account['end_balance'] ?? 0);
                    $currencyTotals[$currency]['difference'] += (float)(($account['end_balance'] ?? 0) - ($account['start_balance'] ?? 0));
                }
                
                // Add professional total rows for each currency
                foreach ($currencyTotals as $currency => $totals) {
                    $sheet->setCellValue('A' . $row, sprintf('TOTAL (%s)', $currency));
                    $sheet->setCellValue('B' . $row, $this->formatCurrency($totals['start_balance']));
                    $sheet->setCellValue('C' . $row, $this->formatCurrency($totals['end_balance']));
                    $sheet->setCellValue('D' . $row, $this->formatCurrency($totals['difference']));
                    $sheet->setCellValue('E' . $row, $currency);
                    
                    // Apply professional total row styling
                    $this->styleTotalRow($sheet, 'A' . $row . ':E' . $row);
                    
                    // Apply professional currency formatting with conditional colors
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, $totals['start_balance'], false);
                    $this->applyCurrencyFormatting($sheet, 'C' . $row, $totals['end_balance'], false);
                    $this->applyCurrencyFormatting($sheet, 'D' . $row, $totals['difference'], false);
                    
                    $row++;
                }
            }
        } else {
            // Try to use the accounts collection directly
            $startDataRow = $row;
            foreach ($this->accounts as $account) {
                $sheet->setCellValue('A' . $row, $account->name);
                $sheet->setCellValue('B' . $row, 'N/A');
                $sheet->setCellValue('C' . $row, 'N/A');
                $sheet->setCellValue('D' . $row, 'N/A');
                $sheet->setCellValue('E' . $row, $account->currency ? $account->currency->symbol : '');
                $row++;
            }
            
            // Add note about totals not being available
            if ($row > $startDataRow + 1) {
                $row++;
                $sheet->setCellValue('A' . $row, 'TOTALS - Not available without structured data');
                $sheet->getStyle('A' . $row . ':E' . $row)->getFont()->setBold(true);
                $sheet->getStyle('A' . $row . ':E' . $row)->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setARGB('FFE6E6');
            }
        }
    }

    /**
     * Populate balance data in the spreadsheet (Income vs Expenses summary)
     */
    private function populateBalanceData($sheet, array $balanceData): void
    {
        $row = 2;
        
        // Debug logging
        app('log')->info('Balance data received for export', ['data' => $balanceData]);
        
        // Handle HTML data by extracting useful information
        if (isset($balanceData['html']) && is_string($balanceData['html'])) {
            $this->parseBalanceDataFromHtml($sheet, $balanceData['html']);
            return;
        }
        
        // If we have operations data with currency breakdown
        if (isset($balanceData['operations']) && is_array($balanceData['operations'])) {
            $operations = $balanceData['operations'];
            $startRow = $row;
            
            foreach ($operations as $currencyId => $data) {
                $currency = $data['currency_symbol'] ?? $data['currency_code'] ?? '';
                
                // Income row with alternating styling
                $sheet->setCellValue('A' . $row, sprintf('Income (%s)', $currency));
                $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($data['in'] ?? 0)));
                $sheet->setCellValue('C' . $row, $currency);
                $isOddRow = ($row - $startRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':C' . $row, $isOddRow);
                $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($data['in'] ?? 0), true);
                // Set default text color for non-currency columns
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'C' . $row);
                $row++;
                
                // Expenses row with alternating styling
                $sheet->setCellValue('A' . $row, sprintf('Expenses (%s)', $currency));
                $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($data['out'] ?? 0)));
                $sheet->setCellValue('C' . $row, $currency);
                $isOddRow = ($row - $startRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':C' . $row, $isOddRow);
                $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($data['out'] ?? 0), false, true);
                // Set default text color for non-currency columns
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'C' . $row);
                $row++;
                
                // Total row using SUM formula with professional styling
                $incomeRow = $row - 2; // Income row is 2 rows above
                $expenseRow = $row - 1; // Expense row is 1 row above
                $netResult = (float)($data['sum'] ?? 0);
                $sheet->setCellValue('A' . $row, sprintf('TOTAL (%s)', $currency));
                $sheet->setCellValue('B' . $row, sprintf('=B%d+B%d', $incomeRow, $expenseRow));
                $sheet->setCellValue('C' . $row, $currency);
                $this->styleTotalRow($sheet, 'A' . $row . ':C' . $row);
                $this->applyCurrencyFormatting($sheet, 'B' . $row, $netResult, false);
                $row++;
                
                // Add separator
                if (count($operations) > 1) {
                    $row++;
                }
            }
        } else {
            // Fallback to old structure
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
                // Calculate TOTAL as formula (Income + Expenses) - fallback version
                $incomeRow = $row - 2; // Income should be 2 rows above
                $expenseRow = $row - 1; // Expenses should be 1 row above
                $sheet->setCellValue('A' . $row, 'TOTAL');
                $sheet->setCellValue('B' . $row, sprintf('=B%d+B%d', $incomeRow, $expenseRow)); // Formula calculation
                $sheet->setCellValue('C' . $row, $balanceData['currency_symbol'] ?? '');
                $sheet->getStyle('B' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                
                $sheet->getStyle('A' . $row . ':C' . $row)->getFont()->setBold(true);
                $row++;
            }
        }
        
        // If no structured data, show data not available
        if ($row === 2) {
            $sheet->setCellValue('A' . $row, 'No balance data available');
            $sheet->setCellValue('B' . $row, 'Check report data source');
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

        // Debug logging for chart data
        app('log')->info('Chart creation started', [
            'has_charts' => isset($this->reportData['charts']),
            'charts_empty' => empty($this->reportData['charts'] ?? []),
            'chart_keys' => array_keys($this->reportData['charts'] ?? []),
            'report_data_keys' => array_keys($this->reportData)
        ]);
        if (!isset($this->reportData['charts']) || empty($this->reportData['charts'])) {
            app('log')->warning('No chart data available for export');
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
     * Create Income vs Expenses chart with monthly breakdown
     */
    private function createIncomeVsExpensesChart(array $balanceData): void
    {
        try {
            // Create new worksheet for the chart
            $chartSheet = $this->spreadsheet->createSheet();
            $chartSheet->setTitle('Income vs Expenses Chart');
            
            // Get monthly breakdown data using AccountTasker
            $monthlyData = $this->getMonthlyIncomeExpensesData();
            
            if (!empty($monthlyData['labels'])) {
                // Create chart data in Chart.js format with monthly breakdown
                $chartData = [
                    'labels' => $monthlyData['labels'], // Month names
                    'datasets' => [
                        [
                            'label' => 'Income',
                            'data' => $monthlyData['income']
                        ],
                        [
                            'label' => 'Expenses', 
                            'data' => $monthlyData['expenses']
                        ]
                    ]
                ];
                
                // Prepare data worksheet
                $dataWorksheet = $this->prepareChartData($chartSheet, $chartData, 'Income vs Expenses by Month');
                
                // Create bar chart
                $chart = $this->createBarChart(
                    $dataWorksheet,
                    $chartData,
                    'Income vs Expenses by Month'
                );
                
                if ($chart !== null) {
                    $chartSheet->addChart($chart);
                }
            } else {
                // No data available - add a message
                $chartSheet->setCellValue('A1', 'Income vs Expenses Chart');
                $chartSheet->setCellValue('A3', 'No monthly chart data available');
                $this->styleHeader($chartSheet, 'A1:C1');
            }
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create income vs expenses chart: %s', $e->getMessage()));
        }
    }

    /**
     * Create default charts for standard reports
     */
    private function createDefaultCharts(): void
    {
        // Create Income vs Expenses chart if balance data is available
        if (isset($this->reportData['balance']) && !empty($this->reportData['balance'])) {
            $this->createIncomeVsExpensesChart($this->reportData['balance']);
        }
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
     * Prepare chart data in worksheet with professional styling
     */
    private function prepareChartData($worksheet, array $chartData, string $title)
    {
        // Add professional chart title
        $worksheet->setCellValue('A1', $title);
        $this->styleSectionHeader($worksheet, 'A1:E1');
        
        $row = 3;
        $dataStartRow = $row;
        
        // Handle different chart data formats
        if (isset($chartData['labels']) && isset($chartData['datasets'])) {
            // ChartJS format
            $labels = $chartData['labels'];
            $datasets = $chartData['datasets'];
            
            // Headers
            $worksheet->setCellValue('A2', 'Label');
            $col = 'B';
            $colCount = 1;
            foreach ($datasets as $dataset) {
                $worksheet->setCellValue($col . '2', $dataset['label'] ?? 'Data');
                $col++;
                $colCount++;
            }
            
            // Style headers
            $headerRange = 'A2:' . chr(65 + $colCount) . '2';
            $this->styleHeader($worksheet, $headerRange);
            
            // Data with alternating row styling
            foreach ($labels as $index => $label) {
                $worksheet->setCellValue('A' . $row, $label);
                $col = 'B';
                foreach ($datasets as $datasetIndex => $dataset) {
                    $value = $dataset['data'][$index] ?? 0;
                    $worksheet->setCellValue($col . $row, $this->formatCurrency((float)$value));
                    
                    // Apply currency formatting with appropriate color
                    $isIncome = isset($dataset['label']) && stripos($dataset['label'], 'income') !== false;
                    $this->applyCurrencyFormatting($worksheet, $col . $row, (float)$value, $isIncome);
                    
                    $col++;
                }
                
                // Apply alternating row styling
                $isOddRow = ($row - $dataStartRow) % 2 === 0;
                $rowRange = 'A' . $row . ':' . chr(65 + $colCount) . $row;
                $this->styleAlternatingRows($worksheet, $rowRange, $isOddRow);
                
                $row++;
            }
        } elseif (is_array($chartData) && !empty($chartData)) {
            // Simple key-value format
            $worksheet->setCellValue('A2', 'Category');
            $worksheet->setCellValue('B2', 'Value');
            $this->styleHeader($worksheet, 'A2:B2');
            
            foreach ($chartData as $key => $value) {
                $worksheet->setCellValue('A' . $row, $key);
                $worksheet->setCellValue('B' . $row, $this->formatCurrency((float)$value));
                
                // Apply currency formatting and alternating row styling
                $this->applyCurrencyFormatting($worksheet, 'B' . $row, (float)$value, false);
                $isOddRow = ($row - $dataStartRow) % 2 === 0;
                $this->styleAlternatingRows($worksheet, 'A' . $row . ':B' . $row, $isOddRow);
                
                $row++;
            }
        }
        
        // Set dynamic column widths
        $this->setDynamicColumnWidths($worksheet, [
            'A' => 20, 'B' => 15, 'C' => 15, 'D' => 15, 'E' => 15
        ]);
        
        return $worksheet;
    }
    
    /**
     * Get chart color from Firefly III palette
     */
    private function getChartColor(int $index): string
    {
        return self::CHART_COLORS[$index % count(self::CHART_COLORS)];
    }
    
    /**
     * Apply professional chart theming
     */
    private function applyChartTheming(Chart $chart): void
    {
        try {
            // Get plot area for styling
            $plotArea = $chart->getPlotArea();
            if ($plotArea) {
                $dataSeries = $plotArea->getPlotGroupByIndex(0)->getDataSeriesCollection();
                
                // Apply Firefly III colors to each data series
                foreach ($dataSeries as $seriesIndex => $series) {
                    $color = $this->getChartColor($seriesIndex);
                    
                    // Apply color theming (this varies by chart type and PhpSpreadsheet version)
                    // Note: PhpSpreadsheet has limited chart customization capabilities
                    // The colors will be more prominent in the data table styling
                }
            }
        } catch (\Exception $e) {
            app('log')->warning('Failed to apply chart theming: ' . $e->getMessage());
        }
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
                        $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$' . $colLetter . '$2',
                        null,
                        1
                    );
                    
                    // Series values
                    $dataSeriesValues[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_NUMBER,
                        $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$' . $colLetter . '$3:$' . $colLetter . '$' . (2 + $rowCount),
                        null,
                        $rowCount
                    );
                }
                
                // Category labels (X-axis)
                $xAxisRange = $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$A$3:$A$' . (2 + $rowCount);
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
                    $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$B$2',
                    null,
                    1
                )
            ];
            
            $dataSeriesValues = [
                new DataSeriesValues(
                    DataSeriesValues::DATASERIES_TYPE_NUMBER,
                    $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$B$3:$B$' . $lastRow,
                    null,
                    $lastRow - 2
                )
            ];
            
            $xAxisLabels = new DataSeriesValues(
                DataSeriesValues::DATASERIES_TYPE_STRING,
                $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$A$3:$A$' . $lastRow,
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
                        $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$' . $colLetter . '$2',
                        null,
                        1
                    );
                    
                    // Series values
                    $dataSeriesValues[] = new DataSeriesValues(
                        DataSeriesValues::DATASERIES_TYPE_NUMBER,
                        $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$' . $colLetter . '$3:$' . $colLetter . '$' . (2 + $rowCount),
                        null,
                        $rowCount
                    );
                }
                
                // Category labels (X-axis)
                $xAxisRange = $this->getQuotedWorksheetName($worksheet->getTitle()) . '!' . '$A$3:$A$' . (2 + $rowCount);
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

    /**
     * Parse account data from HTML response
     */
    private function parseAccountDataFromHtml($sheet, string $html): void
    {
        $row = 2;
        
        // Try to extract data from HTML table (basic approach)
        if (preg_match_all('/<tr.*?>(.*?)<\/tr>/s', $html, $matches)) {
            foreach ($matches[1] as $tableRow) {
                if (preg_match_all('/<td.*?>(.*?)<\/td>/s', $tableRow, $cellMatches)) {
                    $cells = $cellMatches[1];
                    if (count($cells) >= 2) {
                        // Extract account name (remove HTML tags)
                        $accountName = strip_tags($cells[0]);
                        if (!empty(trim($accountName)) && trim($accountName) !== 'Account') {
                            $sheet->setCellValue('A' . $row, trim($accountName));
                            $sheet->setCellValue('B' . $row, 'See web report');
                            $sheet->setCellValue('C' . $row, 'See web report');
                            $sheet->setCellValue('D' . $row, 'See web report');
                            $row++;
                        }
                    }
                }
            }
        }
        
        if ($row === 2) {
            $sheet->setCellValue('A' . $row, 'No account data parsed from HTML');
            $sheet->setCellValue('B' . $row, 'Check data collection process');
        }
    }

    /**
     * Parse balance data from HTML response
     */
    private function parseBalanceDataFromHtml($sheet, string $html): void
    {
        $row = 2;
        
        // Extract balance information from HTML
        if (preg_match('/income.*?([0-9,.]+)/i', strip_tags($html), $matches)) {
            $sheet->setCellValue('A' . $row, 'Income');
            $sheet->setCellValue('B' . $row, $matches[1]);
            $sheet->setCellValue('C' . $row, 'From HTML');
            $row++;
        }
        
        if (preg_match('/expense.*?([0-9,.]+)/i', strip_tags($html), $matches)) {
            $sheet->setCellValue('A' . $row, 'Expenses');
            $sheet->setCellValue('B' . $row, $matches[1]);
            $sheet->setCellValue('C' . $row, 'From HTML');
            $row++;
        }
        
        if ($row === 2) {
            $sheet->setCellValue('A' . $row, 'No balance data parsed from HTML');
            $sheet->setCellValue('B' . $row, 'Check data collection process');
        }
    }

    /**
     * Populate income data in the spreadsheet
     */
    private function populateIncomeData($sheet, array $incomeData): void
    {
        $row = 2;
        
        app('log')->info('Income data received for export', ['data' => $incomeData]);
        
        // Handle HTML data
        if (isset($incomeData['html']) && is_string($incomeData['html'])) {
            $this->parseIncomeDataFromHtml($sheet, $incomeData['html']);
            return;
        }
        
        // Handle structured data with professional styling
        if (isset($incomeData['accounts']) && is_array($incomeData['accounts'])) {
            $startDataRow = $row;
            foreach ($incomeData['accounts'] as $accountKey => $account) {
                $sheet->setCellValue('A' . $row, $account['name'] ?? 'N/A');
                $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($account['sum'] ?? 0)));
                $sheet->setCellValue('C' . $row, $account['currency_symbol'] ?? '');
                
                // Apply professional styling and formatting
                $amount = (float)($account['sum'] ?? 0);
                $this->applyCurrencyFormatting($sheet, 'B' . $row, $amount, true);
                
                // Apply alternating row styling
                $isOddRow = ($row - $startDataRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':C' . $row, $isOddRow);
                
                // Set default text color for non-currency columns
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'C' . $row);
                
                $row++;
            }
            
            // Add totals if we have sums data
            if (isset($incomeData['sums']) && is_array($incomeData['sums'])) {
                $row++; // Empty row separator
                
                // Calculate range for SUM formula - sum all income rows for this currency
                $endDataRow = $row - 2; // Last data row before separator
                $sumFormula = sprintf('=SUM(B%d:B%d)', $startDataRow, $endDataRow);
                
                foreach ($incomeData['sums'] as $currencyId => $sum) {
                    $sheet->setCellValue('A' . $row, sprintf('TOTAL (%s)', $sum['currency_symbol'] ?? $sum['currency_code'] ?? ''));
                    $sheet->setCellValue('B' . $row, $sumFormula); // Use SUM formula instead of hardcoded value
                    $sheet->setCellValue('C' . $row, $sum['currency_symbol'] ?? '');
                    
                    // Apply professional total row styling
                    $this->styleTotalRow($sheet, 'A' . $row . ':C' . $row);
                    
                    // Apply professional currency formatting for income
                    $amount = (float)($sum['sum'] ?? 0);
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, $amount, true);
                    
                    $row++;
                }
            }
        } else {
            $sheet->setCellValue('A' . $row, 'No income data available');
        }
    }

    /**
     * Populate expenses data in the spreadsheet
     */
    private function populateExpensesData($sheet, array $expensesData): void
    {
        $row = 2;
        
        app('log')->info('Expenses data received for export', ['data' => $expensesData]);
        
        // Handle HTML data
        if (isset($expensesData['html']) && is_string($expensesData['html'])) {
            $this->parseExpensesDataFromHtml($sheet, $expensesData['html']);
            return;
        }
        
        // Handle structured data with professional styling
        if (isset($expensesData['accounts']) && is_array($expensesData['accounts'])) {
            $startDataRow = $row;
            foreach ($expensesData['accounts'] as $accountKey => $account) {
                $sheet->setCellValue('A' . $row, $account['name'] ?? 'N/A');
                $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($account['sum'] ?? 0)));
                $sheet->setCellValue('C' . $row, $account['currency_symbol'] ?? '');
                
                // Apply professional styling and formatting
                $amount = (float)($account['sum'] ?? 0);
                $this->applyCurrencyFormatting($sheet, 'B' . $row, $amount, false, true);
                
                // Apply alternating row styling
                $isOddRow = ($row - $startDataRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':C' . $row, $isOddRow);
                
                // Set default text color for non-currency columns
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'C' . $row);
                
                $row++;
            }
            
            // Add totals if we have sums data
            if (isset($expensesData['sums']) && is_array($expensesData['sums'])) {
                $row++; // Empty row separator
                
                // Calculate range for SUM formula - sum all expense rows for this currency
                $endDataRow = $row - 2; // Last data row before separator
                $sumFormula = sprintf('=SUM(B%d:B%d)', $startDataRow, $endDataRow);
                
                foreach ($expensesData['sums'] as $currencyId => $sum) {
                    $sheet->setCellValue('A' . $row, sprintf('TOTAL (%s)', $sum['currency_symbol'] ?? $sum['currency_code'] ?? ''));
                    $sheet->setCellValue('B' . $row, $sumFormula); // Use SUM formula instead of hardcoded value
                    $sheet->setCellValue('C' . $row, $sum['currency_symbol'] ?? '');
                    
                    // Apply professional total row styling
                    $this->styleTotalRow($sheet, 'A' . $row . ':C' . $row);
                    
                    // Apply professional currency formatting for expenses
                    $amount = (float)($sum['sum'] ?? 0);
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, $amount, false, true);
                    
                    $row++;
                }
            }
        } else {
            $sheet->setCellValue('A' . $row, 'No expenses data available');
        }
    }

    /**
     * Parse income data from HTML response
     */
    private function parseIncomeDataFromHtml($sheet, string $html): void
    {
        $this->parseAccountOperationsFromHtml($sheet, $html, 'income');
    }

    /**
     * Parse expenses data from HTML response
     */
    private function parseExpensesDataFromHtml($sheet, string $html): void
    {
        $this->parseAccountOperationsFromHtml($sheet, $html, 'expenses');
    }

    /**
     * Parse account operations data from HTML
     */
    private function parseAccountOperationsFromHtml($sheet, string $html, string $type): void
    {
        $row = 2;
        
        // Try to extract data from HTML table
        if (preg_match_all('/<tr.*?>(.*?)<\/tr>/s', $html, $matches)) {
            foreach ($matches[1] as $tableRow) {
                if (preg_match_all('/<td.*?>(.*?)<\/td>/s', $tableRow, $cellMatches)) {
                    $cells = $cellMatches[1];
                    if (count($cells) >= 2) {
                        $accountName = strip_tags($cells[0]);
                        if (!empty(trim($accountName)) && trim($accountName) !== 'Account' && trim($accountName) !== $type) {
                            $sheet->setCellValue('A' . $row, trim($accountName));
                            $sheet->setCellValue('B' . $row, 'See web report');
                            $sheet->setCellValue('C' . $row, 'N/A');
                            $row++;
                        }
                    }
                }
            }
        }
        
        if ($row === 2) {
            $sheet->setCellValue('A' . $row, sprintf('No %s data parsed from HTML', $type));
        }
    }

    /**
     * Populate categories data in the spreadsheet
     */
    private function populateCategoriesData($sheet, array $categoriesData): void
    {
        $row = 2;
        
        app('log')->info('Categories data received for export', ['data' => $categoriesData]);
        
        // Handle structured data from CategoryReportGenerator with alternating styling
        if (isset($categoriesData['categories']) && is_array($categoriesData['categories'])) {
            $startDataRow = $row;
            
            foreach ($categoriesData['categories'] as $categoryKey => $category) {
                // Handle both 'name' and 'title' fields, and provide proper fallback for no category
                $categoryName = $category['name'] ?? $category['title'] ?? null;
                if (empty($categoryName) || $categoryName === 'No category') {
                    $categoryName = '(No category)';
                }
                
                $sheet->setCellValue('A' . $row, $categoryName);
                $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($category['spent'] ?? 0)));
                $sheet->setCellValue('C' . $row, $this->formatCurrency((float)($category['earned'] ?? 0)));
                $sheet->setCellValue('D' . $row, $this->formatCurrency((float)($category['sum'] ?? 0)));
                $sheet->setCellValue('E' . $row, $category['currency_symbol'] ?? '');
                
                // Apply alternating row styling
                $isOddRow = ($row - $startDataRow) % 2 === 0;
                $this->styleAlternatingRows($sheet, 'A' . $row . ':E' . $row, $isOddRow);
                
                // Apply currency formatting with color coding
                $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($category['spent'] ?? 0), false, true); // Spent (expenses)
                $this->applyCurrencyFormatting($sheet, 'C' . $row, (float)($category['earned'] ?? 0), true); // Earned (income)
                $this->applyCurrencyFormatting($sheet, 'D' . $row, (float)($category['sum'] ?? 0), false); // Net sum
                
                // Set default text color for non-currency columns
                $this->setDefaultTextColor($sheet, 'A' . $row);
                $this->setDefaultTextColor($sheet, 'E' . $row);
                
                $row++;
            }
            
            // Add totals if we have sums data
            if (isset($categoriesData['sums']) && is_array($categoriesData['sums'])) {
                $row++; // Empty separator row
                
                foreach ($categoriesData['sums'] as $currencyId => $sum) {
                    $currency = $sum['currency_symbol'] ?? $sum['currency_code'] ?? '';
                    $sheet->setCellValue('A' . $row, sprintf('TOTAL (%s)', $currency));
                    $sheet->setCellValue('B' . $row, $this->formatCurrency((float)($sum['spent'] ?? 0)));
                    $sheet->setCellValue('C' . $row, $this->formatCurrency((float)($sum['earned'] ?? 0)));
                    $sheet->setCellValue('D' . $row, $this->formatCurrency((float)($sum['sum'] ?? 0)));
                    $sheet->setCellValue('E' . $row, $currency);
                    
                    // Apply professional total row styling
                    $this->styleTotalRow($sheet, 'A' . $row . ':E' . $row);
                    
                    // Apply currency formatting with color coding
                    $this->applyCurrencyFormatting($sheet, 'B' . $row, (float)($sum['spent'] ?? 0), false, true);
                    $this->applyCurrencyFormatting($sheet, 'C' . $row, (float)($sum['earned'] ?? 0), true);
                    $this->applyCurrencyFormatting($sheet, 'D' . $row, (float)($sum['sum'] ?? 0), false);
                    
                    $row++;
                }
            }
        } else {
            $sheet->setCellValue('A' . $row, 'No categories data available');
        }
    }

    /**
     * Get monthly income vs expenses data for chart
     */
    private function getMonthlyIncomeExpensesData(): array
    {
        try {
            $tasker = app(AccountTaskerInterface::class);
            $tasker->setUser($this->user);
            
            $labels = [];
            $incomeData = [];
            $expenseData = [];
            
            // Create monthly intervals from start to end date
            $current = $this->start->copy()->startOfMonth();
            $end = $this->end->copy()->endOfMonth();
            
            while ($current <= $end) {
                $monthStart = $current->copy()->startOfMonth();
                $monthEnd = $current->copy()->endOfMonth();
                
                // Only process months that overlap with our date range
                if ($monthEnd >= $this->start && $monthStart <= $this->end) {
                    $actualStart = $monthStart < $this->start ? $this->start : $monthStart;
                    $actualEnd = $monthEnd > $this->end ? $this->end : $monthEnd;
                    
                    $labels[] = $current->format('M Y');
                    
                    // Get income and expense totals for this month
                    $incomes = $tasker->getIncomeReport($actualStart, $actualEnd, $this->accounts);
                    $expenses = $tasker->getExpenseReport($actualStart, $actualEnd, $this->accounts);
                    
                    // Sum across all currencies (using first currency found for simplicity)
                    $monthlyIncome = 0;
                    $monthlyExpense = 0;
                    
                    foreach ($incomes['sums'] as $currencyData) {
                        $monthlyIncome += (float)($currencyData['sum'] ?? 0);
                    }
                    
                    foreach ($expenses['sums'] as $currencyData) {
                        $monthlyExpense += abs((float)($currencyData['sum'] ?? 0)); // Make expenses positive
                    }
                    
                    $incomeData[] = $monthlyIncome;
                    $expenseData[] = $monthlyExpense;
                }
                
                $current->addMonth();
            }
            
            return [
                'labels' => $labels,
                'income' => $incomeData,
                'expenses' => $expenseData
            ];
            
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to get monthly income/expense data: %s', $e->getMessage()));
            return ['labels' => [], 'income' => [], 'expenses' => []];
        }
    }
    
    /**
     * Get properly quoted worksheet name for Excel formulas
     */
    private function getQuotedWorksheetName(string $worksheetName): string
    {
        // If the worksheet name contains spaces or special characters, wrap in single quotes
        if (preg_match("/[\\s\\-]/", $worksheetName)) {
            return "'" . str_replace("'", "''", $worksheetName) . "'";
        }
        return $worksheetName;
    }
    
    /**
     * Create monthly categories sheet with income and expense tables
     */
    private function createMonthlyCategoriesSheet($sheet): void
    {
        try {
            // Get monthly category data
            $monthlyData = $this->getMonthlyCategoryData();
            
            if (empty($monthlyData['months'])) {
                $sheet->setCellValue('A1', 'Categories');
                $sheet->setCellValue('A3', 'No category data available');
                $this->styleHeader($sheet, 'A1:C1');
                return;
            }
            
            $row = 1;
            
            // Title
            $sheet->setCellValue('A' . $row, 'Category Income and Expense Analysis');
            $this->styleHeader($sheet, 'A' . $row . ':' . chr(65 + count($monthlyData['months'])) . $row);
            $row += 2;
            
            // Income Table
            $sheet->setCellValue('A' . $row, 'INCOME BY CATEGORY');
            $this->styleHeader($sheet, 'A' . $row . ':' . chr(65 + count($monthlyData['months'])) . $row);
            $row++;
            
            // Income table headers
            $sheet->setCellValue('A' . $row, 'Category');
            $col = 'B';
            foreach ($monthlyData['months'] as $month) {
                $sheet->setCellValue($col . $row, $month);
                $col++;
            }
            $sheet->setCellValue($col . $row, 'Total');
            $this->styleHeader($sheet, 'A' . $row . ':' . $col . $row);
            $row++;
            
            // Income data rows with alternating styling
            $incomeStartRow = $row;
            foreach ($monthlyData['categories'] as $categoryName => $categoryData) {
                if ($categoryData['total_income'] != 0) {
                    $sheet->setCellValue('A' . $row, $categoryName);
                    $col = 'B';
                    $startCol = $col;
                    foreach ($monthlyData['months'] as $month) {
                        $amount = $categoryData['months'][$month]['income'] ?? 0;
                        $sheet->setCellValue($col . $row, $this->formatCurrency($amount));
                        $this->applyCurrencyFormatting($sheet, $col . $row, $amount, true);
                        $col++;
                    }
                    // Use SUM formula for row total
                    $endCol = chr(ord($col) - 1); // Previous column
                    $sumFormula = sprintf('=SUM(%s%d:%s%d)', $startCol, $row, $endCol, $row);
                    $sheet->setCellValue($col . $row, $sumFormula);
                    $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                    $sheet->getStyle($col . $row)->getFont()->setBold(true);
                    
                    // Apply alternating row styling
                    $isOddRow = ($row - $incomeStartRow) % 2 === 0;
                    $this->styleAlternatingRows($sheet, 'A' . $row . ':' . $col . $row, $isOddRow);
                    
                    // Set default text color for category name
                    $this->setDefaultTextColor($sheet, 'A' . $row);
                    
                    $row++;
                }
            }
            
            // Income totals row with SUM formulas
            $sheet->setCellValue('A' . $row, 'TOTAL INCOME');
            $col = 'B';
            $endDataRow = $row - 1; // Last data row
            foreach ($monthlyData['months'] as $month) {
                // Create SUM formula for each month column
                $sumFormula = sprintf('=SUM(%s%d:%s%d)', $col, $incomeStartRow, $col, $endDataRow);
                $sheet->setCellValue($col . $row, $sumFormula);
                $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                $col++;
            }
            // Total column - sum all monthly totals
            $lastMonthCol = chr(65 + count($monthlyData['months'])); // Convert to column letter
            $totalFormula = sprintf('=SUM(B%d:%s%d)', $row, $lastMonthCol, $row);
            $sheet->setCellValue($col . $row, $totalFormula);
            $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $this->styleTotalRow($sheet, 'A' . $row . ':' . $col . $row);
            $row += 3;
            
            // Expense Table
            $sheet->setCellValue('A' . $row, 'EXPENSES BY CATEGORY');
            $this->styleHeader($sheet, 'A' . $row . ':' . chr(65 + count($monthlyData['months'])) . $row);
            $row++;
            
            // Expense table headers
            $sheet->setCellValue('A' . $row, 'Category');
            $col = 'B';
            foreach ($monthlyData['months'] as $month) {
                $sheet->setCellValue($col . $row, $month);
                $col++;
            }
            $sheet->setCellValue($col . $row, 'Total');
            $this->styleHeader($sheet, 'A' . $row . ':' . $col . $row);
            $row++;
            
            // Expense data rows with alternating styling
            $expenseStartRow = $row;
            foreach ($monthlyData['categories'] as $categoryName => $categoryData) {
                if ($categoryData['total_expense'] != 0) {
                    $sheet->setCellValue('A' . $row, $categoryName);
                    $col = 'B';
                    $startCol = $col;
                    foreach ($monthlyData['months'] as $month) {
                        $amount = $categoryData['months'][$month]['expense'] ?? 0;
                        $sheet->setCellValue($col . $row, $this->formatCurrency(abs($amount))); // Show as positive
                        $this->applyCurrencyFormatting($sheet, $col . $row, $amount, false, true); // Expenses should be red
                        $col++;
                    }
                    // Use SUM formula for row total
                    $endCol = chr(ord($col) - 1); // Previous column
                    $sumFormula = sprintf('=SUM(%s%d:%s%d)', $startCol, $row, $endCol, $row);
                    $sheet->setCellValue($col . $row, $sumFormula);
                    $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                    $sheet->getStyle($col . $row)->getFont()->setBold(true);
                    
                    // Apply alternating row styling
                    $isOddRow = ($row - $expenseStartRow) % 2 === 0;
                    $this->styleAlternatingRows($sheet, 'A' . $row . ':' . $col . $row, $isOddRow);
                    
                    // Set default text color for category name
                    $this->setDefaultTextColor($sheet, 'A' . $row);
                    
                    $row++;
                }
            }
            
            // Expense totals row with SUM formulas
            $sheet->setCellValue('A' . $row, 'TOTAL EXPENSES');
            $col = 'B';
            $endDataRow = $row - 1; // Last data row
            foreach ($monthlyData['months'] as $month) {
                // Create SUM formula for each month column
                $sumFormula = sprintf('=SUM(%s%d:%s%d)', $col, $expenseStartRow, $col, $endDataRow);
                $sheet->setCellValue($col . $row, $sumFormula);
                $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
                $col++;
            }
            // Total column - sum all monthly totals
            $lastMonthCol = chr(65 + count($monthlyData['months'])); // Convert to column letter
            $totalFormula = sprintf('=SUM(B%d:%s%d)', $row, $lastMonthCol, $row);
            $sheet->setCellValue($col . $row, $totalFormula);
            $sheet->getStyle($col . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $this->styleTotalRow($sheet, 'A' . $row . ':' . $col . $row);
            
            // Set column widths
            $sheet->getColumnDimension('A')->setWidth(25);
            for ($i = 1; $i <= count($monthlyData['months']) + 1; $i++) {
                $sheet->getColumnDimension(chr(65 + $i))->setWidth(12);
            }
            
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to create monthly categories sheet: %s', $e->getMessage()));
            $sheet->setCellValue('A1', 'Categories');
            $sheet->setCellValue('A3', 'Error creating category data: ' . $e->getMessage());
        }
    }
    
    /**
     * Get monthly category data for detailed breakdown
     */
    private function getMonthlyCategoryData(): array
    {
        try {
            $generator = app(CategoryReportGenerator::class);
            $generator->setUser($this->user);
            $generator->setAccounts($this->accounts);
            
            $months = [];
            $categories = [];
            $totals = ['income' => 0, 'expense' => 0];
            
            // Create monthly intervals from start to end date
            $current = $this->start->copy()->startOfMonth();
            $end = $this->end->copy()->endOfMonth();
            
            while ($current <= $end) {
                $monthStart = $current->copy()->startOfMonth();
                $monthEnd = $current->copy()->endOfMonth();
                
                // Only process months that overlap with our date range
                if ($monthEnd >= $this->start && $monthStart <= $this->end) {
                    $actualStart = $monthStart < $this->start ? $this->start : $monthStart;
                    $actualEnd = $monthEnd > $this->end ? $this->end : $monthEnd;
                    
                    $monthLabel = $current->format('M Y');
                    $months[] = $monthLabel;
                    
                    // Get category data for this month
                    $generator->setStart($actualStart);
                    $generator->setEnd($actualEnd);
                    $generator->operations();
                    $monthData = $generator->getReport();
                    
                    if (isset($monthData['categories'])) {
                        foreach ($monthData['categories'] as $categoryKey => $categoryInfo) {
                            $categoryName = $categoryInfo['name'] ?? $categoryInfo['title'] ?? '(No category)';
                            
                            if (!isset($categories[$categoryName])) {
                                $categories[$categoryName] = [
                                    'months' => [],
                                    'total_income' => 0,
                                    'total_expense' => 0
                                ];
                            }
                            
                            $income = max(0, (float)($categoryInfo['earned'] ?? 0));
                            $expense = min(0, (float)($categoryInfo['spent'] ?? 0));
                            
                            $categories[$categoryName]['months'][$monthLabel] = [
                                'income' => $income,
                                'expense' => $expense
                            ];
                            
                            $categories[$categoryName]['total_income'] += $income;
                            $categories[$categoryName]['total_expense'] += abs($expense);
                            
                            $totals['income'] += $income;
                            $totals['expense'] += abs($expense);
                        }
                    }
                }
                
                $current->addMonth();
            }
            
            // Fill in missing months for each category
            foreach ($categories as &$categoryData) {
                foreach ($months as $month) {
                    if (!isset($categoryData['months'][$month])) {
                        $categoryData['months'][$month] = [
                            'income' => 0,
                            'expense' => 0
                        ];
                    }
                }
            }
            
            return [
                'months' => $months,
                'categories' => $categories,
                'totals' => $totals
            ];
            
        } catch (\Exception $e) {
            app('log')->error(sprintf('Failed to get monthly category data: %s', $e->getMessage()));
            return ['months' => [], 'categories' => [], 'totals' => []];
        }
    }
}
