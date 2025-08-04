/*
 * dates.js
 * Copyright (c) 2023 james@firefly-iii.org
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

import {
    addMonths,
    endOfDay,
    endOfMonth,
    endOfQuarter,
    endOfWeek,
    startOfDay,
    startOfMonth,
    startOfQuarter,
    startOfWeek,
    startOfYear,
    subDays,
    subMonths
} from "date-fns";
import format from '../../util/format'
import i18next from "i18next";

export default () => ({
    range: {
        start: null, end: null
    },
    defaultRange: {
        start: null, end: null
    },
    language: 'en_US',

    init() {
        this.range = {
            start: new Date(window.store.get('start')),
            end: new Date(window.store.get('end'))
        };
        this.defaultRange = {
            start: new Date(window.store.get('start')),
            end: new Date(window.store.get('end'))
        };
        this.language = window.store.get('language');
        this.locale = window.store.get('locale');
        this.locale = 'equal' === this.locale ? this.language : this.locale;
        window.__localeId__ = this.language;
        this.buildDateRange();

        window.store.observe('start', (newValue) => {
            this.range.start = new Date(newValue);
        });
        window.store.observe('end', (newValue) => {
            this.range.end = new Date(newValue);
            this.buildDateRange();
        });
    },


    buildDateRange() {
        // console.log('Dates buildDateRange');

        // generate ranges
        let nextRange = this.getNextRange();
        let prevRange = this.getPrevRange();
        let last7 = this.lastDays(7);
        let last30 = this.lastDays(30);
        let mtd = this.mtd();
        let ytd = this.ytd();
        let fyc = this.currentFinancialYear();
        let fyp = this.previousFinancialYear();
        let fy3 = this.thirdLastFinancialYear();

        // set the title:
        let element = document.getElementsByClassName('daterange-holder')[0];
        element.textContent = format(this.range.start) + ' - ' + format(this.range.end);
        element.setAttribute('data-start', format(this.range.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(this.range.end, 'yyyy-MM-dd'));

        // set the current one
        element = document.getElementsByClassName('daterange-current')[0];
        element.textContent = format(this.defaultRange.start) + ' - ' + format(this.defaultRange.end);
        element.setAttribute('data-start', format(this.defaultRange.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(this.defaultRange.end, 'yyyy-MM-dd'));

        // generate next range
        element = document.getElementsByClassName('daterange-next')[0];
        element.textContent = format(nextRange.start) + ' - ' + format(nextRange.end);
        element.setAttribute('data-start', format(nextRange.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(nextRange.end, 'yyyy-MM-dd'));

        // previous range.
        element = document.getElementsByClassName('daterange-prev')[0];
        element.textContent = format(prevRange.start) + ' - ' + format(prevRange.end);
        element.setAttribute('data-start', format(prevRange.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(prevRange.end, 'yyyy-MM-dd'));

        // last 7
        element = document.getElementsByClassName('daterange-7d')[0];
        element.setAttribute('data-start', format(last7.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(last7.end, 'yyyy-MM-dd'));

        // last 30
        element = document.getElementsByClassName('daterange-90d')[0];
        element.setAttribute('data-start', format(last30.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(last30.end, 'yyyy-MM-dd'));

        // MTD
        element = document.getElementsByClassName('daterange-mtd')[0];
        element.setAttribute('data-start', format(mtd.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(mtd.end, 'yyyy-MM-dd'));

        // YTD
        element = document.getElementsByClassName('daterange-ytd')[0];
        element.setAttribute('data-start', format(ytd.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(ytd.end, 'yyyy-MM-dd'));

        // Current Financial Year
        element = document.getElementsByClassName('daterange-fyc')[0];
        element.setAttribute('data-start', format(fyc.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(fyc.end, 'yyyy-MM-dd'));
        element.textContent = i18next.t('firefly.current_financial_year');

        // Previous Financial Year
        element = document.getElementsByClassName('daterange-fyp')[0];
        element.setAttribute('data-start', format(fyp.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(fyp.end, 'yyyy-MM-dd'));
        element.textContent = i18next.t('firefly.previous_financial_year', {year: fyp.end.getFullYear()});

        // Third Last Financial Year
        element = document.getElementsByClassName('daterange-fy3')[0];
        element.setAttribute('data-start', format(fy3.start, 'yyyy-MM-dd'));
        element.setAttribute('data-end', format(fy3.end, 'yyyy-MM-dd'));
        element.textContent = i18next.t('firefly.financial_year_x', {year: fy3.end.getFullYear()});

        // custom range.
        // console.log('MainApp: buildDateRange end');
    },

    getNextRange() {
        let start = startOfMonth(this.range.start);
        let nextMonth = addMonths(start, 1);
        let end = endOfMonth(nextMonth);
        return {start: nextMonth, end: end};
    },

    getPrevRange() {
        let start = startOfMonth(this.range.start);
        let prevMonth = subMonths(start, 1);
        let end = endOfMonth(prevMonth);
        return {start: prevMonth, end: end};
    },

    ytd() {
        let end = new Date;
        let start = startOfYear(this.range.start);
        return {start: start, end: end};
    },

    mtd() {

        let end = new Date;
        let start = startOfMonth(this.range.start);
        return {start: start, end: end};
    },

    lastDays(days) {
        let end = new Date;
        let start = subDays(end, days);
        return {start: start, end: end};
    },

    changeDateRange(e) {
        e.preventDefault();
        // console.log('MainApp: changeDateRange');
        let target = e.currentTarget;

        let start = new Date(target.getAttribute('data-start'));
        let end = new Date(target.getAttribute('data-end'));
        // console.log('MainApp: Change date range', start, end);

        window.store.set('start', start);
        window.store.set('end', end);
        //this.buildDateRange();
        return false;
    },

    currentFinancialYear() {
        // Get fiscal year start preference from store
        const fiscalYearStart = window.store.get('fiscalYearStart') || '01-01';
        const useCustomFiscalYear = window.store.get('customFiscalYear') || false;
        
        let start, end;
        if (useCustomFiscalYear) {
            const [month, day] = fiscalYearStart.split('-');
            start = new Date(new Date().getFullYear(), parseInt(month) - 1, parseInt(day));
            
            // If the fiscal year start is in the future, subtract a year
            if (start > new Date()) {
                start.setFullYear(start.getFullYear() - 1);
            }
            
            end = new Date(start);
            end.setFullYear(end.getFullYear() + 1);
            end.setDate(end.getDate() - 1);
        } else {
            // Calendar year
            start = startOfYear(new Date());
            end = new Date(start.getFullYear(), 11, 31);
        }
        
        return {start: start, end: end};
    },

    previousFinancialYear() {
        const current = this.currentFinancialYear();
        let start = new Date(current.start);
        let end = new Date(current.end);
        
        start.setFullYear(start.getFullYear() - 1);
        end.setFullYear(end.getFullYear() - 1);
        
        return {start: start, end: end};
    },

    thirdLastFinancialYear() {
        const current = this.currentFinancialYear();
        let start = new Date(current.start);
        let end = new Date(current.end);
        
        start.setFullYear(start.getFullYear() - 2);
        end.setFullYear(end.getFullYear() - 2);
        
        return {start: start, end: end};
    },

});
