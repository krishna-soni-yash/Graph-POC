import * as React from 'react';
import type { IPocProps } from './IPocProps';
import ScheduleVariation, { IScheduleVariationRow } from './Graphs/ScheduleVaration';
import styles from './Poc.module.scss';
import ScheduleVariationService, { IProjectMetricsItem, ITaskManagementItem, IWorkLogManagementItem } from '../../../services/ScheduleVariation';

interface IPocState {
  metricItem: IProjectMetricsItem | null;
  selectedGraph: '' | 'scheduleVariation';
  isLoading: boolean;
  loadError?: string;
}

export default class Poc extends React.Component<IPocProps, IPocState> {
  private readonly _scheduleVariationService: ScheduleVariationService;

  public constructor(props: IPocProps) {
    super(props);

    this.state = {
      metricItem: null,
      selectedGraph: '',
      isLoading: false,
      loadError: undefined
    };

    this._scheduleVariationService = new ScheduleVariationService(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this._loadScheduleVariationData();
  }

  private _loadScheduleVariationData = async (): Promise<void> => {
    this.setState({ isLoading: true, loadError: undefined });

    try {
      const metricItem = await this._scheduleVariationService.getScheduleVariationMetric();
      this.setState({ metricItem, isLoading: false });
    } catch (error) {
      console.error('Failed to fetch Schedule Variation data from ProjectMetrics list.', error);
      this.setState({ isLoading: false, loadError: 'Unable to fetch Schedule Variation data.' });
    }
  };

  private _onGraphChange = async (event: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
    const selectedGraph = event.target.value as '' | 'scheduleVariation';
    this.setState({ selectedGraph });

    if (selectedGraph === 'scheduleVariation') {
      await this._loadScheduleVariationData();
    }
  };

  private _getMinimumDate(dateValues: Array<string | undefined>): string | undefined {
    const validDates = dateValues
      .filter((value): value is string => !!value)
      .map((value) => new Date(value))
      .filter((date) => !isNaN(date.getTime()));

    if (validDates.length === 0) {
      return undefined;
    }

    return new Date(Math.min(...validDates.map((date) => date.getTime()))).toISOString();
  }

  private _getMaximumDate(dateValues: Array<string | undefined>): string | undefined {
    const validDates = dateValues
      .filter((value): value is string => !!value)
      .map((value) => new Date(value))
      .filter((date) => !isNaN(date.getTime()));

    if (validDates.length === 0) {
      return undefined;
    }

    return new Date(Math.max(...validDates.map((date) => date.getTime()))).toISOString();
  }

  private _calculateScheduleVariation(
    plannedEndDate?: string,
    actualEndDate?: string,
    plannedDuration?: number
  ): number | undefined {
    if (!plannedEndDate || !actualEndDate || plannedDuration === undefined || plannedDuration === 0) {
      return undefined;
    }

    const plannedEnd = new Date(plannedEndDate);
    const actualEnd = new Date(actualEndDate);

    if (isNaN(plannedEnd.getTime()) || isNaN(actualEnd.getTime())) {
      return undefined;
    }

    const millisecondsPerDay = 1000 * 60 * 60 * 24;
    const dateDifferenceInDays = (actualEnd.getTime() - plannedEnd.getTime()) / millisecondsPerDay;
    return (dateDifferenceInDays * 100) / plannedDuration;
  }

  private _buildScheduleVariationRows(metricItem: IProjectMetricsItem | null): IScheduleVariationRow[] {
    if (!metricItem) {
      return [];
    }

    const relatedWorkLogs = metricItem.RelatedWorkLogs ?? [];
    const relatedTasks = metricItem.RelatedTasks ?? [];

    return relatedWorkLogs.map((workLog: IWorkLogManagementItem) => {
      const workLogId = Number(workLog.ID);
      const workLogNo = workLog.WorkItemNo !== undefined ? Number(workLog.WorkItemNo) : undefined;
      const matchingTasks = relatedTasks.filter(
        (task: ITaskManagementItem) => {
          const taskLookupId = task.WorkItemNoId !== undefined ? Number(task.WorkItemNoId) : undefined;
          const taskWorkItemNo = task.WorkItemNo !== undefined ? Number(task.WorkItemNo) : undefined;
          return taskLookupId === workLogId || taskWorkItemNo === workLogId || taskWorkItemNo === workLogNo;
        }
      );

      const actualStartDate = this._getMinimumDate(matchingTasks.map((task) => task.ActualStartDate));
      const actualEndDate = this._getMaximumDate(matchingTasks.map((task) => task.ActualEndDate));
      const scheduleVaration = this._calculateScheduleVariation(
        workLog.PlannedEndDate,
        actualEndDate,
        workLog.PlannedDuration
      );
      const squaredDeviationOfScheduleVariation =
        typeof metricItem.SquaredDeviationOfScheduleVariation === 'number'
          ? metricItem.SquaredDeviationOfScheduleVariation
          : undefined;

      return {
        WorkItemNo: workLog.WorkItemNo ?? workLog.ID,
        PlannedStartDate: workLog.PlannedStartDate,
        PlannedEndDate: workLog.PlannedEndDate,
        PlannedDuration: workLog.PlannedDuration,
        ActualStartDate: actualStartDate,
        ActualEndDate: actualEndDate,
        ScheduleVaration: scheduleVaration,
        SquaredDeviationOfScheduleVariation: squaredDeviationOfScheduleVariation,
        Variance: metricItem.Variance,
        SquareRootOfVariance: metricItem.SquareRootOfVariance,
        Goal: metricItem.Goal,
        USL: metricItem.USL,
        LSL: metricItem.LSL
      };
    });
  }

  public render(): React.ReactElement<IPocProps> {
    const {
      hasTeamsContext
    } = this.props;
    const { metricItem, selectedGraph, isLoading, loadError } = this.state;
    const scheduleVariationRows = this._buildScheduleVariationRows(metricItem);

    return (
      <section className={`${styles.defects} ${hasTeamsContext ? styles.teams : ''}`}>
        <div
          style={{
            maxWidth: '960px',
            margin: '0 auto',
            borderRadius: '14px',
            padding: '18px',
            background: 'linear-gradient(145deg, #f8fbff 0%, #eef6ff 45%, #f6fdf9 100%)',
            border: '1px solid #dbeafe',
            boxShadow: '0 10px 24px rgba(15, 23, 42, 0.08)'
          }}
        >
          <div
            style={{
              display: 'flex',
              flexWrap: 'wrap',
              alignItems: 'center',
              justifyContent: 'space-between',
              gap: '14px',
              marginBottom: '16px'
            }}
          >
            <div>
              <h2
                style={{
                  margin: 0,
                  fontSize: '22px',
                  letterSpacing: '0.2px',
                  color: '#0f172a'
                }}
              >
                Project Metrics Dashboard
              </h2>
              <p
                style={{
                  margin: '4px 0 0',
                  fontSize: '13px',
                  color: '#334155'
                }}
              >
                Select a graph to visualize insights from SharePoint list data.
              </p>
            </div>

            <div style={{ minWidth: '230px' }}>
              <label
                htmlFor="graph-selector"
                style={{
                  display: 'block',
                  marginBottom: '6px',
                  fontSize: '12px',
                  fontWeight: 700,
                  color: '#1e293b',
                  textTransform: 'uppercase',
                  letterSpacing: '0.7px'
                }}
              >
                Graph
              </label>
              <select
                id="graph-selector"
                value={selectedGraph}
                onChange={this._onGraphChange}
                style={{
                  width: '100%',
                  height: '40px',
                  border: '1px solid #93c5fd',
                  borderRadius: '10px',
                  padding: '0 12px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  color: '#0f172a',
                  cursor: 'pointer'
                }}
              >
                <option value="">Select Graph</option>
                <option value="scheduleVariation">Schedule Variation</option>
              </select>
            </div>
          </div>

          {selectedGraph === 'scheduleVariation' ? (
            isLoading ? (
              <div
                style={{
                  border: '1px dashed #93c5fd',
                  borderRadius: '12px',
                  padding: '28px 20px',
                  textAlign: 'center',
                  color: '#475569',
                  backgroundColor: 'rgba(255, 255, 255, 0.7)'
                }}
              >
                Loading Schedule Variation data...
              </div>
            ) : loadError ? (
              <div
                style={{
                  border: '1px solid #fecaca',
                  borderRadius: '12px',
                  padding: '28px 20px',
                  textAlign: 'center',
                  color: '#991b1b',
                  backgroundColor: '#fef2f2'
                }}
              >
                {loadError}
              </div>
            ) : (
              <ScheduleVariation
                title={(metricItem?.Metrics as string) || 'Schedule Variation'}
                rows={scheduleVariationRows}
              />
            )
          ) : (
            <div
              style={{
                border: '1px dashed #93c5fd',
                borderRadius: '12px',
                padding: '28px 20px',
                textAlign: 'center',
                color: '#475569',
                backgroundColor: 'rgba(255, 255, 255, 0.7)'
              }}
            >
              Choose <strong>Schedule Variation</strong> from the dropdown to display the graph.
            </div>
          )}
        </div>
      </section>
    );
  }
}
