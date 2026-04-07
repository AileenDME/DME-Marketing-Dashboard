import React, { useState, useEffect } from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, BarChart, Bar } from 'recharts';
import { parseISO, format, subDays, startOfMonth, endOfMonth, startOfQuarter, endOfQuarter, startOfYear, endOfYear, isSameDay, isWithinInterval } from 'date-fns';
import ExcelDataService from './services/ExcelDataService';
import './Dashboard.css';

const Dashboard = () => {
  const [excelData, setExcelData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [lastUpdated, setLastUpdated] = useState(new Date());

  // Leads Generated State
  const [leadsDateRange, setLeadsDateRange] = useState('last30');
  const [leadsCustomStart, setLeadsCustomStart] = useState('');
  const [leadsCustomEnd, setLeadsCustomEnd] = useState('');
  const [leadsGranularity, setLeadsGranularity] = useState('daily');
  const [leadsShowComparison, setLeadsShowComparison] = useState(false);

  // Leads Velocity State
  const [velocityDateRange, setVelocityDateRange] = useState('last30');
  const [velocityCustomStart, setVelocityCustomStart] = useState('');
  const [velocityCustomEnd, setVelocityCustomEnd] = useState('');
  const [velocityGranularity, setVelocityGranularity] = useState('daily');
  const [velocityShowComparison, setVelocityShowComparison] = useState(false);

  // Branding State (DA, Traffic, LinkedIn)
  const [brandingDateRange, setBrandingDateRange] = useState('last30');
  const [brandingCustomStart, setBrandingCustomStart] = useState('');
  const [brandingCustomEnd, setBrandingCustomEnd] = useState('');
  const [daGranularity, setDaGranularity] = useState('daily');
  const [trafficGranularity, setTrafficGranularity] = useState('daily');
  const [linkedinGranularity, setLinkedinGranularity] = useState('daily');
  const [daCompareLastYear, setDaCompareLastYear] = useState(false);
  const [trafficCompareLastYear, setTrafficCompareLastYear] = useState(false);
  const [linkedinCompareLastYear, setLinkedinCompareLastYear] = useState(false);

  // BD Enablement State
  const [bdDateRange, setBdDateRange] = useState('last30');
  const [bdCustomStart, setBdCustomStart] = useState('');
  const [bdCustomEnd, setBdCustomEnd] = useState('');

  // Fetch data on mount
  useEffect(() => {
    const fetchData = async () => {
      try {
        setLoading(true);
        const data = await ExcelDataService.fetchFromSharePoint();
        setExcelData(data);
        setLastUpdated(new Date());
      } catch (err) {
        console.error('Error fetching data:', err);
        setError(err.message || 'Failed to load data');
        // For demo, load sample data
        setExcelData(await ExcelDataService.loadSampleData());
      } finally {
        setLoading(false);
      }
    };

    fetchData();
    // Refresh every hour
    const interval = setInterval(fetchData, 3600000);
    return () => clearInterval(interval);
  }, []);

  const getDateRange = (rangeType, customStart = null, customEnd = null) => {
    const today = new Date();
    const endDate = customEnd ? parseISO(customEnd) : today;
    let startDate;

    switch (rangeType) {
      case 'last30':
        startDate = subDays(today, 30);
        break;
      case 'mtd':
        startDate = startOfMonth(today);
        break;
      case 'qtd':
        startDate = startOfQuarter(today);
        break;
      case 'ytd':
        startDate = startOfYear(today);
        break;
      case 'custom':
        startDate = customStart ? parseISO(customStart) : subDays(today, 30);
        break;
      default:
        startDate = subDays(today, 30);
    }

    return { startDate, endDate };
  };

  const filterDataByDateRange = (data, dateRange, customStart, customEnd) => {
    const { startDate, endDate } = getDateRange(dateRange, customStart, customEnd);
    return data.filter(item => {
      const itemDate = parseISO(item.date);
      return isWithinInterval(itemDate, { start: startDate, end: endDate });
    });
  };

  const aggregateByGranularity = (data, granularity) => {
    if (granularity === 'daily') return data;

    const grouped = {};
    data.forEach(item => {
      const date = parseISO(item.date);
      let key;
      if (granularity === 'weekly') {
        const weekStart = new Date(date);
        weekStart.setDate(date.getDate() - date.getDay());
        key = format(weekStart, 'yyyy-MM-dd');
      } else if (granularity === 'monthly') {
        key = format(date, 'yyyy-MM-01');
      }

      if (!grouped[key]) {
        grouped[key] = { date: key, values: [] };
      }
      grouped[key].values.push(item.value);
    });

    return Object.values(grouped).map(g => ({
      date: g.date,
      value: Math.round(g.values.reduce((a, b) => a + b, 0) / g.values.length)
    }));
  };

  const calculatePreviousPeriod = (data, dateRange, customStart, customEnd) => {
    const { startDate, endDate } = getDateRange(dateRange, customStart, customEnd);
    const periodLength = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));
    const prevStart = new Date(startDate.getTime() - periodLength * 24 * 60 * 60 * 1000);
    const prevEnd = new Date(startDate.getTime());

    return data.filter(item => {
      const itemDate = parseISO(item.date);
      return isWithinInterval(itemDate, { start: prevStart, end: prevEnd });
    });
  };

  const calculateComparison = (currentData, previousData, type = 'total') => {
    if (!previousData.length) return 0;

    let currentVal, prevVal;

    if (type === 'total') {
      currentVal = currentData.reduce((sum, d) => sum + d.value, 0);
      prevVal = previousData.reduce((sum, d) => sum + d.value, 0);
    } else if (type === 'average') {
      currentVal = currentData.reduce((sum, d) => sum + d.value, 0) / currentData.length;
      prevVal = previousData.reduce((sum, d) => sum + d.value, 0) / previousData.length;
    }

    const change = ((currentVal - prevVal) / prevVal) * 100;
    return {
      value: currentVal,
      change: Math.round(change),
      isPositive: change >= 0
    };
  };

  if (loading) return <div className="dashboard-container"><div className="loading">Loading dashboard...</div></div>;
  if (error && !excelData) return <div className="dashboard-container"><div className="error">{error}</div></div>;

  return (
    <div className="dashboard-container">
      <div className="dashboard-header">
        <h1>Marketing KPI Dashboard</h1>
        <p>Last updated {format(lastUpdated, 'MMMM d, yyyy HH:mm')}</p>
      </div>

      {excelData && (
        <>
          {/* LEADS GENERATED */}
          <section className="kpi-section">
            <h2>Leads Generated</h2>
            <div className="section-controls">
              <div className="control-group">
                <label>Date Range</label>
                <select value={leadsDateRange} onChange={(e) => setLeadsDateRange(e.target.value)}>
                  <option value="last30">Last 30 days</option>
                  <option value="mtd">Month to date</option>
                  <option value="qtd">Quarter to date</option>
                  <option value="ytd">Year to date</option>
                  <option value="custom">Custom range</option>
                </select>
                {leadsDateRange === 'custom' && (
                  <>
                    <input type="date" value={leadsCustomStart} onChange={(e) => setLeadsCustomStart(e.target.value)} placeholder="From" />
                    <input type="date" value={leadsCustomEnd} onChange={(e) => setLeadsCustomEnd(e.target.value)} placeholder="To" />
                  </>
                )}
              </div>
              <div className="control-group">
                <label>Granularity</label>
                <select value={leadsGranularity} onChange={(e) => setLeadsGranularity(e.target.value)}>
                  <option value="daily">Daily</option>
                  <option value="weekly">Weekly</option>
                  <option value="monthly">Monthly</option>
                </select>
              </div>
              <div className="control-group checkbox">
                <label>
                  <input type="checkbox" checked={leadsShowComparison} onChange={(e) => setLeadsShowComparison(e.target.checked)} />
                  Compare with previous period
                </label>
              </div>
            </div>

            {excelData.leadsGenerated && (
              <LeadsGeneratedCard
                data={excelData.leadsGenerated}
                dateRange={leadsDateRange}
                customStart={leadsCustomStart}
                customEnd={leadsCustomEnd}
                granularity={leadsGranularity}
                showComparison={leadsShowComparison}
                filterDataByDateRange={filterDataByDateRange}
                aggregateByGranularity={aggregateByGranularity}
                calculatePreviousPeriod={calculatePreviousPeriod}
                calculateComparison={calculateComparison}
              />
            )}
            <p className="source-text"><a href={excelData.sharePointLink} target="_blank" rel="noopener noreferrer">Source: Leads Generated sheet</a></p>
          </section>

          {/* LEADS VELOCITY */}
          <section className="kpi-section">
            <h2>Leads Velocity</h2>
            <div className="section-controls">
              <div className="control-group">
                <label>Date Range</label>
                <select value={velocityDateRange} onChange={(e) => setVelocityDateRange(e.target.value)}>
                  <option value="last30">Last 30 days</option>
                  <option value="mtd">Month to date</option>
                  <option value="qtd">Quarter to date</option>
                  <option value="ytd">Year to date</option>
                  <option value="custom">Custom range</option>
                </select>
                {velocityDateRange === 'custom' && (
                  <>
                    <input type="date" value={velocityCustomStart} onChange={(e) => setVelocityCustomStart(e.target.value)} placeholder="From" />
                    <input type="date" value={velocityCustomEnd} onChange={(e) => setVelocityCustomEnd(e.target.value)} placeholder="To" />
                  </>
                )}
              </div>
              <div className="control-group">
                <label>Granularity</label>
                <select value={velocityGranularity} onChange={(e) => setVelocityGranularity(e.target.value)}>
                  <option value="daily">Daily</option>
                  <option value="weekly">Weekly</option>
                  <option value="monthly">Monthly</option>
                </select>
              </div>
              <div className="control-group checkbox">
                <label>
                  <input type="checkbox" checked={velocityShowComparison} onChange={(e) => setVelocityShowComparison(e.target.checked)} />
                  Compare with previous period
                </label>
              </div>
            </div>

            {excelData.leadsVelocity && (
              <LeadsVelocityCard
                data={excelData.leadsVelocity}
                dateRange={velocityDateRange}
                customStart={velocityCustomStart}
                customEnd={velocityCustomEnd}
                granularity={velocityGranularity}
                showComparison={velocityShowComparison}
                filterDataByDateRange={filterDataByDateRange}
                aggregateByGranularity={aggregateByGranularity}
                calculatePreviousPeriod={calculatePreviousPeriod}
                calculateComparison={calculateComparison}
              />
            )}
            <p className="source-text"><a href={excelData.sharePointLink} target="_blank" rel="noopener noreferrer">Source: Leads Velocity sheet</a></p>
          </section>

          {/* BRANDING SECTION */}
          <section className="kpi-section">
            <h2>Branding</h2>
            <div className="section-controls">
              <div className="control-group">
                <label>Date Range</label>
                <select value={brandingDateRange} onChange={(e) => setBrandingDateRange(e.target.value)}>
                  <option value="last30">Last 30 days</option>
                  <option value="mtd">Month to date</option>
                  <option value="qtd">Quarter to date</option>
                  <option value="ytd">Year to date</option>
                  <option value="custom">Custom range</option>
                </select>
                {brandingDateRange === 'custom' && (
                  <>
                    <input type="date" value={brandingCustomStart} onChange={(e) => setBrandingCustomStart(e.target.value)} placeholder="From" />
                    <input type="date" value={brandingCustomEnd} onChange={(e) => setBrandingCustomEnd(e.target.value)} placeholder="To" />
                  </>
                )}
              </div>
            </div>

            <div className="branding-grid">
              {excelData.domainAuthority && (
                <BrandingMetric
                  title="Domain Authority"
                  data={excelData.domainAuthority}
                  dateRange={brandingDateRange}
                  customStart={brandingCustomStart}
                  customEnd={brandingCustomEnd}
                  granularity={daGranularity}
                  setGranularity={setDaGranularity}
                  compareLastYear={daCompareLastYear}
                  setCompareLastYear={setDaCompareLastYear}
                  filterDataByDateRange={filterDataByDateRange}
                  aggregateByGranularity={aggregateByGranularity}
                  calculateComparison={calculateComparison}
                  sheetName="Domain Authority"
                />
              )}
              {excelData.usTraffic && (
                <BrandingMetric
                  title="US Organic Traffic"
                  data={excelData.usTraffic}
                  dateRange={brandingDateRange}
                  customStart={brandingCustomStart}
                  customEnd={brandingCustomEnd}
                  granularity={trafficGranularity}
                  setGranularity={setTrafficGranularity}
                  compareLastYear={trafficCompareLastYear}
                  setCompareLastYear={setTrafficCompareLastYear}
                  filterDataByDateRange={filterDataByDateRange}
                  aggregateByGranularity={aggregateByGranularity}
                  calculateComparison={calculateComparison}
                  sheetName="US Traffic"
                />
              )}
              {excelData.linkedinFollowers && (
                <BrandingMetric
                  title="LinkedIn Followers (US)"
                  data={excelData.linkedinFollowers}
                  dateRange={brandingDateRange}
                  customStart={brandingCustomStart}
                  customEnd={brandingCustomEnd}
                  granularity={linkedinGranularity}
                  setGranularity={setLinkedinGranularity}
                  compareLastYear={linkedinCompareLastYear}
                  setCompareLastYear={setLinkedinCompareLastYear}
                  filterDataByDateRange={filterDataByDateRange}
                  aggregateByGranularity={aggregateByGranularity}
                  calculateComparison={calculateComparison}
                  sheetName="LinkedIn Followers"
                />
              )}
            </div>
          </section>

          {/* BD ENABLEMENT */}
          <section className="kpi-section">
            <h2>BD Enablement</h2>
            <div className="section-controls">
              <div className="control-group">
                <label>Date Range</label>
                <select value={bdDateRange} onChange={(e) => setBdDateRange(e.target.value)}>
                  <option value="last30">Last 30 days</option>
                  <option value="mtd">Month to date</option>
                  <option value="qtd">Quarter to date</option>
                  <option value="ytd">Year to date</option>
                  <option value="custom">Custom range</option>
                </select>
                {bdDateRange === 'custom' && (
                  <>
                    <input type="date" value={bdCustomStart} onChange={(e) => setBdCustomStart(e.target.value)} placeholder="From" />
                    <input type="date" value={bdCustomEnd} onChange={(e) => setBdCustomEnd(e.target.value)} placeholder="To" />
                  </>
                )}
              </div>
            </div>

            <div className="bd-grid">
              {excelData.salesDeck && (
                <BDMetricCard
                  title="Sales Decks"
                  data={excelData.salesDeck}
                  dateRange={bdDateRange}
                  customStart={bdCustomStart}
                  customEnd={bdCustomEnd}
                  filterDataByDateRange={filterDataByDateRange}
                  sheetName="Sales Deck Created"
                />
              )}
              {excelData.rfpResponses && (
                <BDMetricCard
                  title="RFP Responses"
                  data={excelData.rfpResponses}
                  dateRange={bdDateRange}
                  customStart={bdCustomStart}
                  customEnd={bdCustomEnd}
                  filterDataByDateRange={filterDataByDateRange}
                  sheetName="RFP Responses"
                />
              )}
            </div>
          </section>

          {/* PROJECTS */}
          {excelData.projects && (
            <section className="kpi-section projects-section">
              <h2>Projects</h2>
              <div className="projects-grid">
                {excelData.projects.map((project, idx) => (
                  <ProjectCard key={idx} project={project} />
                ))}
              </div>
              <p className="source-text"><a href={excelData.sharePointLink} target="_blank" rel="noopener noreferrer">Source: Projects sheet</a></p>
            </section>
          )}
        </>
      )}
    </div>
  );
};

// Component for Leads Generated Card
const LeadsGeneratedCard = ({
  data,
  dateRange,
  customStart,
  customEnd,
  granularity,
  showComparison,
  filterDataByDateRange,
  aggregateByGranularity,
  calculatePreviousPeriod,
  calculateComparison
}) => {
  const filteredData = filterDataByDateRange(data, dateRange, customStart, customEnd);
  const aggregatedData = aggregateByGranularity(filteredData, granularity);
  const previousData = showComparison ? aggregateByGranularity(calculatePreviousPeriod(data, dateRange, customStart, customEnd), granularity) : [];
  const { value: totalLeads, change, isPositive } = calculateComparison(filteredData, calculatePreviousPeriod(data, dateRange, customStart, customEnd), 'total');

  const chartData = aggregatedData.map((item, idx) => ({
    date: format(parseISO(item.date), 'MMM dd'),
    current: item.value,
    ...(showComparison && previousData[idx] && { previous: previousData[idx].value })
  }));

  return (
    <div className="metric-card-container">
      <div className="metric-display">
        <div className="metric-value">{Math.round(totalLeads)}</div>
        <div className={`metric-change ${isPositive ? 'positive' : 'negative'}`}>
          {isPositive ? '↑' : '↓'} {Math.abs(change)}% vs previous period
        </div>
      </div>
      <div className="chart-container">
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="date" />
            <YAxis />
            <Tooltip />
            {showComparison && <Legend />}
            <Line type="monotone" dataKey="current" stroke="#185FA5" name="Current Period" dot={{ r: 4 }} />
            {showComparison && <Line type="monotone" dataKey="previous" stroke="#ccc" name="Previous Period" dot={{ r: 4 }} />}
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
};

// Component for Leads Velocity Card
const LeadsVelocityCard = ({
  data,
  dateRange,
  customStart,
  customEnd,
  granularity,
  showComparison,
  filterDataByDateRange,
  aggregateByGranularity,
  calculatePreviousPeriod,
  calculateComparison
}) => {
  const filteredData = filterDataByDateRange(data, dateRange, customStart, customEnd);
  const aggregatedData = aggregateByGranularity(filteredData, granularity);
  const previousData = showComparison ? aggregateByGranularity(calculatePreviousPeriod(data, dateRange, customStart, customEnd), granularity) : [];
  const { value: avgVelocity, change, isPositive } = calculateComparison(filteredData, calculatePreviousPeriod(data, dateRange, customStart, customEnd), 'average');

  const chartData = aggregatedData.map((item, idx) => ({
    date: format(parseISO(item.date), 'MMM dd'),
    current: item.value.toFixed(2),
    ...(showComparison && previousData[idx] && { previous: previousData[idx].value.toFixed(2) })
  }));

  return (
    <div className="metric-card-container">
      <div className="metric-display">
        <div className="metric-value">{(avgVelocity * 100).toFixed(1)}%</div>
        <div className={`metric-change ${isPositive ? 'positive' : 'negative'}`}>
          {isPositive ? '↑' : '↓'} {Math.abs(change)}% vs previous period
        </div>
      </div>
      <div className="chart-container">
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="date" />
            <YAxis />
            <Tooltip />
            {showComparison && <Legend />}
            <Line type="monotone" dataKey="current" stroke="#10b981" name="Current Period" dot={{ r: 4 }} />
            {showComparison && <Line type="monotone" dataKey="previous" stroke="#ccc" name="Previous Period" dot={{ r: 4 }} />}
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
};

// Component for Branding Metrics
const BrandingMetric = ({
  title,
  data,
  dateRange,
  customStart,
  customEnd,
  granularity,
  setGranularity,
  compareLastYear,
  setCompareLastYear,
  filterDataByDateRange,
  aggregateByGranularity,
  calculateComparison,
  sheetName
}) => {
  const filteredData = filterDataByDateRange(data, dateRange, customStart, customEnd);
  const aggregatedData = aggregateByGranularity(filteredData, granularity);
  const latestValue = aggregatedData.length > 0 ? aggregatedData[aggregatedData.length - 1].value : 0;

  const chartData = aggregatedData.map(item => ({
    date: format(parseISO(item.date), 'MMM dd'),
    value: item.value
  }));

  return (
    <div className="branding-metric">
      <div className="metric-header">
        <h3>{title}</h3>
        <div className="metric-controls">
          <select value={granularity} onChange={(e) => setGranularity(e.target.value)} className="small-select">
            <option value="daily">Daily</option>
            <option value="weekly">Weekly</option>
            <option value="monthly">Monthly</option>
          </select>
          <label className="checkbox-label">
            <input type="checkbox" checked={compareLastYear} onChange={(e) => setCompareLastYear(e.target.checked)} />
            vs last month
          </label>
        </div>
      </div>
      <div className="metric-value">{latestValue}</div>
      <div className="chart-container">
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="date" />
            <YAxis />
            <Tooltip />
            <Line type="monotone" dataKey="value" stroke="#f59e0b" dot={{ r: 3 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>
      <p className="source-text"><a href="#" onClick={(e) => e.preventDefault()}>Source: {sheetName}</a></p>
    </div>
  );
};

// Component for BD Metrics
const BDMetricCard = ({
  title,
  data,
  dateRange,
  customStart,
  customEnd,
  filterDataByDateRange,
  sheetName
}) => {
  const filteredData = filterDataByDateRange(data, dateRange, customStart, customEnd);
  const totalValue = filteredData.reduce((sum, d) => sum + d.value, 0);

  return (
    <div className="bd-metric">
      <h3>{title}</h3>
      <div className="metric-value-large">{Math.round(totalValue)}</div>
      <p className="source-text"><a href="#" onClick={(e) => e.preventDefault()}>Source: {sheetName}</a></p>
    </div>
  );
};

// Component for Projects
const ProjectCard = ({ project }) => {
  const statusColor = {
    'On Track': '#10b981',
    'At Risk': '#f59e0b',
    'Blocked': '#ef4444',
    'On Hold': '#6b7280',
    'Completed': '#3b82f6'
  };

  return (
    <div className="project-card">
      <div className="project-header">
        <h4>{project.name}</h4>
        <span className="project-status" style={{ color: statusColor[project.status] || '#6b7280' }}>
          {project.status}
        </span>
      </div>
      <div className="progress-bar">
        <div className="progress-fill" style={{ width: `${project.percentage * 100}%`, backgroundColor: statusColor[project.status] || '#6b7280' }}></div>
      </div>
      <div className="progress-text">{Math.round(project.percentage * 100)}%</div>
    </div>
  );
};

export default Dashboard;
