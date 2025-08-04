<?php

/**
 * ReportExportGenerator.php
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
use FireflyIII\Support\Export\SpreadsheetExportService;
use FireflyIII\Support\Export\ReportDataCollector;
use FireflyIII\User;
use Illuminate\Support\Collection;

/**
 * Class ReportExportGenerator
 * 
 * Orchestrates the export of complete Firefly III reports including charts and tables
 */
class ReportExportGenerator
{
    private User $user;
    private string $reportType;
    private Collection $accounts;
    private Collection $budgets;
    private Collection $categories;
    private Collection $tags;
    private Collection $expense;
    private Carbon $start;
    private Carbon $end;
    private SpreadsheetExportService $spreadsheetService;
    private ReportDataCollector $dataCollector;

    public function __construct()
    {
        $this->accounts = new Collection();
        $this->budgets = new Collection();
        $this->categories = new Collection();
        $this->tags = new Collection();
        $this->expense = new Collection();
        $this->spreadsheetService = app(SpreadsheetExportService::class);
        $this->dataCollector = app(ReportDataCollector::class);
    }

    /**
     * Generate the complete report export
     * 
     * @throws FireflyException
     */
    public function export(): string
    {
        // Initialize data collector with report parameters
        $this->dataCollector->setUser($this->user);
        $this->dataCollector->setReportType($this->reportType);
        $this->dataCollector->setDateRange($this->start, $this->end);
        $this->dataCollector->setAccounts($this->accounts);
        
        // Set additional parameters based on report type
        if ($this->budgets->count() > 0) {
            $this->dataCollector->setBudgets($this->budgets);
        }
        
        if ($this->categories->count() > 0) {
            $this->dataCollector->setCategories($this->categories);
        }
        
        if ($this->tags->count() > 0) {
            $this->dataCollector->setTags($this->tags);
        }
        
        if ($this->expense->count() > 0) {
            $this->dataCollector->setExpenseAccounts($this->expense);
        }

        // Collect all report data
        $reportData = $this->dataCollector->collectReportData();

        // Initialize spreadsheet service with report parameters
        $this->spreadsheetService->setUser($this->user);
        $this->spreadsheetService->setReportType($this->reportType);
        $this->spreadsheetService->setDateRange($this->start, $this->end);
        $this->spreadsheetService->setAccounts($this->accounts);
        $this->spreadsheetService->setBudgets($this->budgets);
        $this->spreadsheetService->setCategories($this->categories);
        $this->spreadsheetService->setTags($this->tags);
        $this->spreadsheetService->setExpenseAccounts($this->expense);
        
        // Pass the collected data to spreadsheet service
        $this->spreadsheetService->setReportData($reportData);

        // Generate and return the Excel file path
        return $this->spreadsheetService->generateSpreadsheet();
    }

    public function setUser(User $user): void
    {
        $this->user = $user;
    }

    public function setReportType(string $reportType): void
    {
        $this->reportType = $reportType;
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

    public function setExpenseAccounts(Collection $expense): void
    {
        $this->expense = $expense;
    }

    public function setDateRange(Carbon $start, Carbon $end): void
    {
        $this->start = $start;
        $this->end = $end;
    }
}