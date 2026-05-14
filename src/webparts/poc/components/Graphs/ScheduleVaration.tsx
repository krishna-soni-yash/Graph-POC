import * as React from 'react';

export interface IScheduleVariationRow {
	WorkItemNo: number;
	PlannedStartDate?: string;
	PlannedEndDate?: string;
	PlannedDuration?: number;
	ActualStartDate?: string;
	ActualEndDate?: string;
	ScheduleVaration?: number;
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
	return (
		<div style={containerStyle}>
			<h3 style={headerStyle}>{title}</h3>
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
							<th style={thStyle}>Goal</th>
							<th style={thStyle}>USL</th>
							<th style={thStyle}>LSL</th>
						</tr>
					</thead>
					<tbody>
						{rows.length === 0 ? (
							<tr>
								<td colSpan={10} style={emptyStyle}>
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
