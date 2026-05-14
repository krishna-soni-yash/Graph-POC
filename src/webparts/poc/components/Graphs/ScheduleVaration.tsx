import * as React from 'react';

export interface IScheduleVariationRow {
	WorkItemNo: number;
	PlannedStartDate?: string;
	PlannedEndDate?: string;
	PlannedDuration?: number;
	ActualStartDate?: string;
	ActualEndDate?: string;
	ScheduleVaration?: number;
	SquaredDeviationOfScheduleVariation?: number;
	Variance?: number;
	SquareRootOfVariance?: number;
	Goal?: number | string;
	USL?: number | string;
	LSL?: number | string;
}

export interface IScheduleVariationProps {
	title?: string;
	rows?: IScheduleVariationRow[];
}

const containerStyle: React.CSSProperties = {
	border: '1px solid #e5e7eb',
	borderRadius: '10px',
	padding: '16px',
	background: '#ffffff'
};

const headerStyle: React.CSSProperties = {
	margin: 0,
	marginBottom: '14px',
	fontSize: '18px',
	fontWeight: 700,
	color: '#111827'
};

const tableWrapperStyle: React.CSSProperties = {
	overflowX: 'auto',
	border: '1px solid #dbeafe',
	borderRadius: '10px'
};

const chartWrapperStyle: React.CSSProperties = {
	marginBottom: '16px',
	border: '1px solid #dbeafe',
	borderRadius: '10px',
	padding: '12px',
	background: '#f8fbff'
};

const chartTitleStyle: React.CSSProperties = {
	margin: '0 0 8px 0',
	fontSize: '13px',
	fontWeight: 700,
	color: '#1e3a8a'
};

const tableStyle: React.CSSProperties = {
	width: '100%',
	borderCollapse: 'collapse',
	minWidth: '980px',
	backgroundColor: '#ffffff'
};

const thStyle: React.CSSProperties = {
	padding: '10px 12px',
	textAlign: 'left',
	fontSize: '12px',
	letterSpacing: '0.4px',
	textTransform: 'uppercase',
	backgroundColor: '#eff6ff',
	color: '#1e3a8a',
	borderBottom: '1px solid #bfdbfe'
};

const tdStyle: React.CSSProperties = {
	padding: '10px 12px',
	fontSize: '13px',
	color: '#0f172a',
	borderBottom: '1px solid #e2e8f0'
};

const mutedValueStyle: React.CSSProperties = {
	color: '#64748b'
};

const emptyStyle: React.CSSProperties = {
	padding: '18px',
	fontSize: '13px',
	color: '#64748b',
	textAlign: 'center'
};

const formatDate = (dateValue?: string): string => {
	if (!dateValue) {
		return '-';
	}

	const parsedDate = new Date(dateValue);
	if (isNaN(parsedDate.getTime())) {
		return '-';
	}

	return parsedDate.toLocaleDateString('en-GB');
};

const formatNumber = (value?: number | string): string => {
	if (value === undefined || value === null || value === '') {
		return '-';
	}

	const numericValue = typeof value === 'number' ? value : Number(value);
	if (isNaN(numericValue)) {
		return String(value);
	}

	return numericValue.toFixed(2);
};

const ScheduleVariation: React.FC<IScheduleVariationProps> = ({
	title = 'Schedule Variation',
	rows = []
}) => {
	const chartRows = rows.filter((row) => typeof row.ScheduleVaration === 'number');
	const chartWidth = 900;
	const chartHeight = 260;
	const chartPadding = { top: 20, right: 28, bottom: 52, left: 58 };
	const plotWidth = chartWidth - chartPadding.left - chartPadding.right;
	const plotHeight = chartHeight - chartPadding.top - chartPadding.bottom;
	const yValues = chartRows.map((row) => row.ScheduleVaration as number);
	const minY = yValues.length > 0 ? Math.min(...yValues) : 0;
	const maxY = yValues.length > 0 ? Math.max(...yValues) : 0;
	const yRange = maxY - minY === 0 ? 1 : maxY - minY;
	const yTickCount: number = 5;
	const yTicks: Array<{ value: number; y: number }> = [];
	for (let index = 0; index < yTickCount; index++) {
		const ratio = yTickCount === 1 ? 0 : index / (yTickCount - 1);
		const value = maxY - ratio * yRange;
		const y = chartPadding.top + ratio * plotHeight;
		yTicks.push({ value, y });
	}

	const points = chartRows.map((row, index) => {
		const x = chartPadding.left + (chartRows.length <= 1 ? plotWidth / 2 : (index * plotWidth) / (chartRows.length - 1));
		const y = chartPadding.top + ((maxY - (row.ScheduleVaration as number)) / yRange) * plotHeight;
		return { x, y, row };
	});

	const linePath = points
		.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x} ${point.y}`)
		.join(' ');

	return (
		<div style={containerStyle}>
			<h3 style={headerStyle}>{title}</h3>
			<div style={chartWrapperStyle}>
				<p style={chartTitleStyle}>Line Chart: X = WorkItemNo, Y = Schedule Variation</p>
				{points.length === 0 ? (
					<div style={emptyStyle}>No chart data available for Schedule Variation.</div>
				) : (
					<svg viewBox={`0 0 ${chartWidth} ${chartHeight}`} width="100%" height="260" role="img" aria-label="Work item number versus schedule variation line chart">
						<line x1={chartPadding.left} y1={chartPadding.top} x2={chartPadding.left} y2={chartHeight - chartPadding.bottom} stroke="#94a3b8" strokeWidth="1" />
						<line x1={chartPadding.left} y1={chartHeight - chartPadding.bottom} x2={chartWidth - chartPadding.right} y2={chartHeight - chartPadding.bottom} stroke="#94a3b8" strokeWidth="1" />

						{yTicks.map((tick: { value: number; y: number }) => (
							<g key={`y-tick-${tick.y}`}>
								<line
									x1={chartPadding.left}
									y1={tick.y}
									x2={chartWidth - chartPadding.right}
									y2={tick.y}
									stroke="#e2e8f0"
									strokeWidth="1"
									strokeDasharray="4 4"
								/>
								<text x={chartPadding.left - 46} y={tick.y + 4} fill="#475569" fontSize="11">
									{tick.value.toFixed(2)}
								</text>
							</g>
						))}
						<text x={chartWidth / 2 - 70} y={chartHeight - 12} fill="#334155" fontSize="11">X-axis: WorkItemNo</text>
						<text transform={`translate(14 ${chartHeight / 2 + 30}) rotate(-90)`} fill="#334155" fontSize="11">Y-axis: Schedule Variation</text>

						<path d={linePath} fill="none" stroke="#2563eb" strokeWidth="2" />

						{points.map((point) => (
							<g key={`chart-point-${point.row.WorkItemNo}`}>
								<circle cx={point.x} cy={point.y} r="4" fill="#1d4ed8">
									<title>{`${point.row.WorkItemNo}-${(point.row.ScheduleVaration as number).toFixed(2)}`}</title>
								</circle>
								<text x={point.x} y={chartHeight - chartPadding.bottom + 14} textAnchor="middle" fill="#334155" fontSize="10">
									{point.row.WorkItemNo}
								</text>
							</g>
						))}
					</svg>
				)}
			</div>
			<div style={tableWrapperStyle}>
				<table style={tableStyle}>
					<thead>
						<tr>
							<th style={thStyle}>WorkItemNo</th>
							<th style={thStyle}>PlannedStartDate</th>
							<th style={thStyle}>PlannedEndDate</th>
							<th style={thStyle}>PlannedDuration</th>
							<th style={thStyle}>ActualStartDate</th>
							<th style={thStyle}>ActualEndDate</th>
							<th style={thStyle}>ScheduleVaration</th>
							<th style={thStyle}>SquaredDeviationOfScheduleVariation</th>
							<th style={thStyle}>Variance</th>
							<th style={thStyle}>SquareRootOfVariance</th>
							<th style={thStyle}>Goal</th>
							<th style={thStyle}>USL</th>
							<th style={thStyle}>LSL</th>
						</tr>
					</thead>
					<tbody>
						{rows.length === 0 ? (
							<tr>
								<td colSpan={13} style={emptyStyle}>
									No rows available for Schedule Variation.
								</td>
							</tr>
						) : (
							rows.map((row) => (
								<tr key={row.WorkItemNo}>
									<td style={tdStyle}>{row.WorkItemNo}</td>
									<td style={tdStyle}>{formatDate(row.PlannedStartDate)}</td>
									<td style={tdStyle}>{formatDate(row.PlannedEndDate)}</td>
									<td style={tdStyle}>{formatNumber(row.PlannedDuration)}</td>
									<td style={tdStyle}>{formatDate(row.ActualStartDate)}</td>
									<td style={tdStyle}>{formatDate(row.ActualEndDate)}</td>
									<td style={tdStyle}>{formatNumber(row.ScheduleVaration)}</td>
									<td style={tdStyle}>{formatNumber(row.SquaredDeviationOfScheduleVariation)}</td>
									<td style={tdStyle}>{formatNumber(row.Variance)}</td>
									<td style={tdStyle}>{formatNumber(row.SquareRootOfVariance)}</td>
									<td style={tdStyle}><span style={mutedValueStyle}>{formatNumber(row.Goal)}</span></td>
									<td style={tdStyle}><span style={mutedValueStyle}>{formatNumber(row.USL)}</span></td>
									<td style={tdStyle}><span style={mutedValueStyle}>{formatNumber(row.LSL)}</span></td>
								</tr>
							))
						)}
					</tbody>
				</table>
			</div>
		</div>
	);
};

export default ScheduleVariation;
