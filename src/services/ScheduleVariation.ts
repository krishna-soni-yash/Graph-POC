import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IProjectMetricsItem {
	Metrics?: string;
	Goal?: number | string;
	USL?: number | string;
	LSL?: number | string;
	RelatedWorkLogs?: IWorkLogManagementItem[];
	RelatedTasks?: ITaskManagementItem[];
	FirstActualStartDate?: string;
	LastActualEndDate?: string;
	LastPlannedEndDate?: string;
	PlannedDuration?: number;
	ScheduleVariation?: number;
	SquaredDeviationOfScheduleVariation?: number;
	SumOfSquaredDeviation?: number;
	Variance?: number;
	SquareRootOfVariance?: number;
	[key: string]: unknown;
}

export interface IWorkLogManagementItem {
	ID: number;
	WorkItemNo?: number;
	PlannedStartDate?: string;
	PlannedEndDate?: string;
	PlannedDuration?: number;
	[key: string]: unknown;
}

export interface ITaskManagementItem {
	ID: number;
	WorkItemNoId?: number;
	WorkItemNo?: number;
	Title?: string;
	ActualStartDate?: string;
	ActualEndDate?: string;
	[key: string]: unknown;
}

const LIST_NAME = 'ProjectMetrics';
const WORK_LOG_LIST_NAME = 'WorkLogManagement';
const TASK_MANAGEMENT_LIST_NAME = 'TaskManagement';
const TARGET_METRIC = 'Schedule Variation';

const escapeODataValue = (value: string): string => value.replace(/'/g, "''");

export default class ScheduleVariationService {
	private readonly _sp: SPFI;

	public constructor(context: WebPartContext) {
		this._sp = spfi().using(SPFx(context));
	}

	public async getScheduleVariationMetric(): Promise<IProjectMetricsItem | null> {
		const filterValue = escapeODataValue(TARGET_METRIC);
		const items = await this._sp.web.lists
			.getByTitle(LIST_NAME)
			.items.select('*')
			.filter(`Metrics eq '${filterValue}'`)
			.top(1)();

		if (items.length === 0) {
			return null;
		}

		const rawMetricItem = items[0] as { [key: string]: unknown };
		const metricItem: IProjectMetricsItem = {
			...rawMetricItem,
			Metrics: this._getStringField(rawMetricItem, 'Metrics'),
			Goal: this._getMetricField(rawMetricItem, 'Goal'),
			USL: this._getMetricField(rawMetricItem, 'USL'),
			LSL: this._getMetricField(rawMetricItem, 'LSL')
		};
		const { previousMonthStartIso, currentMonthStartIso } = this._getPreviousMonthBoundaries();
		const workLogItems = await this._sp.web.lists
			.getByTitle(WORK_LOG_LIST_NAME)
			.items.select('ID, WorkItemNo, PlannedStartDate, PlannedEndDate')
			.filter(
				`(Status eq 'Completed' or Status eq 'Closed') and ProjectType eq 'dev' and Modified ge datetime'${previousMonthStartIso}' and Modified lt datetime'${currentMonthStartIso}'`
			)();

		metricItem.RelatedWorkLogs = (workLogItems as IWorkLogManagementItem[]).map((item) => ({
			...item,
			PlannedDuration: this._calculatePlannedDuration(item.PlannedStartDate, item.PlannedEndDate)
		}));

		if (metricItem.RelatedWorkLogs.length > 0) {
			const workLogIds = metricItem.RelatedWorkLogs
				.map((item) => item.ID)
				.filter((id): id is number => typeof id === 'number');

			if (workLogIds.length > 0) {
				const taskFilter = workLogIds.map((id) => `WorkItemNoId eq ${id}`).join(' or ');
				const taskItems = await this._sp.web.lists
					.getByTitle(TASK_MANAGEMENT_LIST_NAME)
					.items.select('ID', 'WorkItemNoId', 'ActualStartDate', 'ActualEndDate')
					.filter(`(${taskFilter})`)();

				metricItem.RelatedTasks = taskItems as ITaskManagementItem[];
				this._populateScheduleVariation(metricItem);
			}
		}
        console.log('Fetched Schedule Variation metric item with related work logs and tasks:', metricItem);
		return metricItem;
	}

	private _getPreviousMonthBoundaries(): { previousMonthStartIso: string; currentMonthStartIso: string } {
		const now = new Date();
		const currentMonthStartUtc = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, 0, 0));
		const previousMonthStartUtc = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1, 0, 0, 0));

		return {
			previousMonthStartIso: previousMonthStartUtc.toISOString(),
			currentMonthStartIso: currentMonthStartUtc.toISOString()
		};
	}

	private _calculatePlannedDuration(plannedStartDate?: string, plannedEndDate?: string): number | undefined {
		if (!plannedStartDate || !plannedEndDate) {
			return undefined;
		}

		const startDate = new Date(plannedStartDate);
		const endDate = new Date(plannedEndDate);

		if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
			return undefined;
		}

		const millisecondsPerDay = 1000 * 60 * 60 * 24;
		return ((endDate.getTime() - startDate.getTime()) / millisecondsPerDay) + 1;
	}

	private _populateScheduleVariation(metricItem: IProjectMetricsItem): void {
		const tasks = metricItem.RelatedTasks ?? [];
		const workLogs = metricItem.RelatedWorkLogs ?? [];

		const firstActualStartDate = this._getMinDate(tasks.map((item) => item.ActualStartDate));
		const lastActualEndDate = this._getMaxDate(tasks.map((item) => item.ActualEndDate));
		const firstPlannedStartDate = this._getMinDate(workLogs.map((item) => item.PlannedStartDate));
		const lastPlannedEndDate = this._getMaxDate(workLogs.map((item) => item.PlannedEndDate));

		if (!lastActualEndDate || !firstPlannedStartDate || !lastPlannedEndDate) {
			return;
		}

		const plannedDuration = this._calculatePlannedDuration(
			firstPlannedStartDate.toISOString(),
			lastPlannedEndDate.toISOString()
		);
		const scheduleVariations: number[] = [];

		for (let i = 0; i < workLogs.length; i++) {
			const workLog = workLogs[i];
			const workLogId = Number(workLog.ID);
			const workLogNo = workLog.WorkItemNo !== undefined ? Number(workLog.WorkItemNo) : undefined;
			const matchingTasks = tasks.filter((task) => {
				const taskLookupId = task.WorkItemNoId !== undefined ? Number(task.WorkItemNoId) : undefined;
				const taskWorkItemNo = task.WorkItemNo !== undefined ? Number(task.WorkItemNo) : undefined;

				return taskLookupId === workLogId || taskWorkItemNo === workLogId || taskWorkItemNo === workLogNo;
			});

			const workLogActualEndDate = this._getMaxDate(matchingTasks.map((item) => item.ActualEndDate));
			if (!workLogActualEndDate || !workLog.PlannedEndDate) {
				continue;
			}

			const workLogPlannedDuration =
				workLog.PlannedDuration ?? this._calculatePlannedDuration(workLog.PlannedStartDate, workLog.PlannedEndDate);
			if (workLogPlannedDuration === undefined || workLogPlannedDuration === 0) {
				continue;
			}

			const workLogPlannedEndDate = new Date(workLog.PlannedEndDate);
			if (isNaN(workLogPlannedEndDate.getTime())) {
				continue;
			}

			const millisecondsPerDay = 1000 * 60 * 60 * 24;
			const endDateDifferenceInDays =
				(workLogActualEndDate.getTime() - workLogPlannedEndDate.getTime()) / millisecondsPerDay;

			scheduleVariations.push((endDateDifferenceInDays * 100) / workLogPlannedDuration);
		}

		metricItem.FirstActualStartDate = firstActualStartDate?.toISOString();
		metricItem.LastActualEndDate = lastActualEndDate.toISOString();
		metricItem.LastPlannedEndDate = lastPlannedEndDate.toISOString();
		metricItem.PlannedDuration = plannedDuration;
		if (scheduleVariations.length > 0) {
			const meanScheduleVariation =
				scheduleVariations.reduce((total, value) => total + value, 0) / scheduleVariations.length;
			const squaredDeviations = scheduleVariations.map((value) => {
				const deviation = value - meanScheduleVariation;
				return deviation * deviation;
			});
			const sumOfSquaredDeviation = squaredDeviations.reduce((total, value) => total + value, 0);
			const squaredDeviationOfScheduleVariation = sumOfSquaredDeviation / scheduleVariations.length;
			const varianceDenominator = meanScheduleVariation - 1;
			const variance =
				varianceDenominator !== 0
					? Math.abs(sumOfSquaredDeviation / varianceDenominator)
					: undefined;
			const squareRootOfVariance = variance !== undefined ? Math.sqrt(variance) : undefined;

			metricItem.ScheduleVariation = meanScheduleVariation;
			metricItem.SquaredDeviationOfScheduleVariation = squaredDeviationOfScheduleVariation;
			metricItem.SumOfSquaredDeviation = sumOfSquaredDeviation;
			metricItem.Variance = variance;
			metricItem.SquareRootOfVariance = squareRootOfVariance;
		} else {
			metricItem.ScheduleVariation = undefined;
			metricItem.SquaredDeviationOfScheduleVariation = undefined;
			metricItem.SumOfSquaredDeviation = undefined;
			metricItem.Variance = undefined;
			metricItem.SquareRootOfVariance = undefined;
		}
	}

	private _getMinDate(dateValues: Array<string | undefined>): Date | undefined {
		const validDates = dateValues
			.filter((value): value is string => !!value)
			.map((value) => new Date(value))
			.filter((date) => !isNaN(date.getTime()));

		if (validDates.length === 0) {
			return undefined;
		}

		return new Date(Math.min(...validDates.map((date) => date.getTime())));
	}

	private _getMaxDate(dateValues: Array<string | undefined>): Date | undefined {
		const validDates = dateValues
			.filter((value): value is string => !!value)
			.map((value) => new Date(value))
			.filter((date) => !isNaN(date.getTime()));

		if (validDates.length === 0) {
			return undefined;
		}

		return new Date(Math.max(...validDates.map((date) => date.getTime())));
	}

	private _getStringField(item: { [key: string]: unknown }, fieldName: string): string | undefined {
		const fieldValue = this._getFieldValue(item, fieldName);
		if (fieldValue === undefined || fieldValue === null) {
			return undefined;
		}

		return String(fieldValue);
	}

	private _getMetricField(item: { [key: string]: unknown }, fieldName: string): number | string | undefined {
		const fieldValue = this._getFieldValue(item, fieldName);
		if (fieldValue === undefined || fieldValue === null || fieldValue === '') {
			return undefined;
		}

		if (typeof fieldValue === 'number') {
			return isNaN(fieldValue) ? undefined : fieldValue;
		}

		if (typeof fieldValue === 'string') {
			const trimmedValue = fieldValue.trim();
			if (trimmedValue === '') {
				return undefined;
			}

			const numericValue = Number(trimmedValue);
			return isNaN(numericValue) ? trimmedValue : numericValue;
		}

		if (typeof fieldValue === 'object') {
			const objectValue = fieldValue as { [key: string]: unknown };
			const candidateKeys = ['LookupValue', 'Title', 'Label', 'Value'];
			for (let i = 0; i < candidateKeys.length; i++) {
				const key = candidateKeys[i];
				const candidate = objectValue[key];
				if (candidate !== undefined && candidate !== null && candidate !== '') {
					const candidateText = String(candidate).trim();
					if (candidateText !== '') {
						const numericValue = Number(candidateText);
						return isNaN(numericValue) ? candidateText : numericValue;
					}
				}
			}
		}

		return String(fieldValue);
	}

	private _getFieldValue(item: { [key: string]: unknown }, fieldName: string): unknown {
		if (Object.prototype.hasOwnProperty.call(item, fieldName)) {
			return item[fieldName];
		}

		const normalizedRequestedName = fieldName.replace(/_/g, '').toLowerCase();
		const keys = Object.keys(item);
		let matchedKey: string | undefined;
		for (let i = 0; i < keys.length; i++) {
			const key: string = keys[i];
			const normalizedKey = key.replace(/_/g, '').toLowerCase();
			if (normalizedKey === normalizedRequestedName || normalizedKey.indexOf(normalizedRequestedName) > -1) {
				matchedKey = key;
				break;
			}
		}

		return matchedKey ? item[matchedKey] : undefined;
	}
}
