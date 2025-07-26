<?php

/**
 * ReportExportController.php
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

namespace FireflyIII\Http\Controllers;

use Carbon\Carbon;
use FireflyIII\Exceptions\FireflyException;
use FireflyIII\Support\Export\ReportExportGenerator;
use Illuminate\Http\JsonResponse;
use Illuminate\Http\Request;
use Illuminate\Http\Response;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Log;

/**
 * Class ReportExportController
 * 
 * Handles exporting complete reports to Excel format
 */
class ReportExportController extends Controller
{
    /**
     * ReportExportController constructor.
     */
    public function __construct()
    {
        parent::__construct();

        $this->middleware(
            function ($request, $next) {
                app('view')->share('title', (string) trans('firefly.report_export'));
                app('view')->share('mainTitleIcon', 'fa-download');

                return $next($request);
            }
        );
    }

    /**
     * Export a default report
     */
    public function defaultReport(Collection $accounts, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('default', $accounts, new Collection(), new Collection(), new Collection(), new Collection(), $start, $end);
    }

    /**
     * Export an audit report
     */
    public function auditReport(Collection $accounts, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('audit', $accounts, new Collection(), new Collection(), new Collection(), new Collection(), $start, $end);
    }

    /**
     * Export a budget report
     */
    public function budgetReport(Collection $accounts, Collection $budgets, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('budget', $accounts, $budgets, new Collection(), new Collection(), new Collection(), $start, $end);
    }

    /**
     * Export a category report
     */
    public function categoryReport(Collection $accounts, Collection $categories, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('category', $accounts, new Collection(), $categories, new Collection(), new Collection(), $start, $end);
    }

    /**
     * Export a tag report
     */
    public function tagReport(Collection $accounts, Collection $tags, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('tag', $accounts, new Collection(), new Collection(), $tags, new Collection(), $start, $end);
    }

    /**
     * Export a double report
     */
    public function doubleReport(Collection $accounts, Collection $expense, Carbon $start, Carbon $end, Request $request)
    {
        return $this->generateExport('double', $accounts, new Collection(), new Collection(), new Collection(), $expense, $start, $end);
    }

    /**
     * Generate and return the report export
     */
    private function generateExport(
        string $reportType,
        Collection $accounts,
        Collection $budgets,
        Collection $categories,
        Collection $tags,
        Collection $expense,
        Carbon $start,
        Carbon $end
    ) {
        try {
            // Validate date range
            if ($end < $start) {
                return response()->json([
                    'error' => trans('firefly.end_after_start_date')
                ], 400);
            }

            // Validate accounts
            if ($accounts->count() === 0) {
                return response()->json([
                    'error' => trans('firefly.select_at_least_one_account')
                ], 400);
            }

            // Create and configure the export generator
            $generator = app(ReportExportGenerator::class);
            $generator->setUser(auth()->user());
            $generator->setReportType($reportType);
            $generator->setAccounts($accounts);
            $generator->setBudgets($budgets);
            $generator->setCategories($categories);
            $generator->setTags($tags);
            $generator->setExpenseAccounts($expense);
            $generator->setDateRange($start, $end);

            // Generate the export
            $filePath = $generator->export();

            // Check if file was created successfully
            if (!file_exists($filePath)) {
                throw new FireflyException('Export file was not created');
            }

            // Read file content
            $content = file_get_contents($filePath);
            if (false === $content) {
                throw new FireflyException('Could not read export file');
            }

            // Determine filename for download
            $filename = basename($filePath);
            $quoted = sprintf('"%s"', addcslashes($filename, '"\\'));

            // Create response using the same pattern as existing export controller
            /** @var \Illuminate\Http\Response $response */
            $response = response($content);
            $response
                ->header('Content-Description', 'File Transfer')
                ->header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                ->header('Content-Disposition', 'attachment; filename=' . $quoted)
                ->header('Content-Transfer-Encoding', 'binary')
                ->header('Connection', 'Keep-Alive')
                ->header('Expires', '0')
                ->header('Cache-Control', 'must-revalidate, post-check=0, pre-check=0')
                ->header('Pragma', 'public')
                ->header('Content-Length', (string) strlen($content));

            // Clean up temporary file
            if (file_exists($filePath)) {
                unlink($filePath);
            }

            return $response;

        } catch (FireflyException $e) {
            Log::error(sprintf('Report export failed: %s', $e->getMessage()));
            Log::error('Report export trace: ' . $e->getTraceAsString());
            
            return response()->json([
                'error' => sprintf('Export failed: %s', $e->getMessage())
            ], 500);
            
        } catch (\Exception $e) {
            Log::error(sprintf('Unexpected error during report export: %s', $e->getMessage()));
            Log::error('Report export trace: ' . $e->getTraceAsString());
            
            return response()->json([
                'error' => sprintf('Export error: %s', $e->getMessage())
            ], 500);
        }
    }

    /**
     * Handle AJAX export status check
     */
    public function exportStatus(Request $request): JsonResponse
    {
        // This can be used for progress tracking in the future
        return response()->json([
            'status' => 'ready',
            'message' => 'Export service is ready'
        ]);
    }
}