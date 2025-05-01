import { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

// Main App Component
const App = () => {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        const fileData = await window.fs.readFile('reportstellarsoiree.xlsx');
        const workbook = XLSX.read(fileData, { cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        setData(jsonData);
        setLoading(false);
      } catch (err) {
        console.error('Error loading data:', err);
        setError('Failed to load data. Please try again.');
        setLoading(false);
      }
    };

    loadData();
  }, []);

  const renderTabContent = () => {
    if (loading) return <div className="text-center p-10">Loading data...</div>;
    if (error) return <div className="text-center p-10 text-red-500">{error}</div>;

    switch (activeTab) {
      case 'dashboard':
        return <Dashboard data={data} />;
      case 'data':
        return <DataTable data={data} />;
      case 'testimoni':
        return <Testimonials data={data} />;
      default:
        return <Dashboard data={data} />;
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-blue-600 text-white p-4 shadow-md">
        <h1 className="text-2xl font-bold text-center">Stellar Soiree RSVP System</h1>
      </header>

      <nav className="bg-white shadow-sm">
        <div className="container mx-auto px-4">
          <ul className="flex space-x-4 overflow-x-auto">
            <li>
              <button
                onClick={() => setActiveTab('dashboard')}
                className={`px-4 py-3 font-medium border-b-2 ${
                  activeTab === 'dashboard'
                    ? 'border-blue-600 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-blue-600'
                }`}
              >
                Dashboard
              </button>
            </li>
            <li>
              <button
                onClick={() => setActiveTab('data')}
                className={`px-4 py-3 font-medium border-b-2 ${
                  activeTab === 'data'
                    ? 'border-blue-600 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-blue-600'
                }`}
              >
                All Data
              </button>
            </li>
            <li>
              <button
                onClick={() => setActiveTab('testimoni')}
                className={`px-4 py-3 font-medium border-b-2 ${
                  activeTab === 'testimoni'
                    ? 'border-blue-600 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-blue-600'
                }`}
              >
                Testimonials
              </button>
            </li>
          </ul>
        </div>
      </nav>

      <main className="container mx-auto px-4 py-6">
        {renderTabContent()}
      </main>
    </div>
  );
};

// Dashboard Component
const Dashboard = ({ data }) => {
  // Prepare data for charts
  const getRsvpStatusCounts = () => {
    const counts = {
      'Hadir': 0,
      'No Respon': 0,
      'Tidak Hadir': 0
    };

    data.forEach(row => {
      const status = row['RSVP'] || 'No Respon';
      counts[status] = (counts[status] || 0) + 1;
    });

    return Object.keys(counts).map(key => ({
      name: key,
      value: counts[key]
    }));
  };

  // Get RSVP counts by class
  const getRsvpByClass = () => {
    const classCounts = {};
    
    // Get all unique classes
    const allClasses = new Set();
    data.forEach(row => {
      if (row['Keterangan 1']) {
        allClasses.add(row['Keterangan 1']);
      }
    });
    
    // Group by class and count RSVP statuses
    Array.from(allClasses).forEach(className => {
      const classData = data.filter(row => row['Keterangan 1'] === className);
      
      const rsvpCounts = {
        'Hadir': 0,
        'No Respon': 0,
        'Tidak Hadir': 0
      };
      
      classData.forEach(row => {
        const status = row['RSVP'] || 'No Respon';
        rsvpCounts[status]++;
      });
      
      classCounts[className] = rsvpCounts;
    });
    
    // Prepare data for the chart
    // Get top 10 classes with most students for readability
    const topClasses = Object.keys(classCounts)
      .filter(c => c && c.startsWith('12'))
      .sort((a, b) => {
        const totalA = Object.values(classCounts[a]).reduce((sum, val) => sum + val, 0);
        const totalB = Object.values(classCounts[b]).reduce((sum, val) => sum + val, 0);
        return totalB - totalA;
      })
      .slice(0, 10);
    
    return topClasses.map(className => ({
      name: className,
      Hadir: classCounts[className]['Hadir'],
      'No Respon': classCounts[className]['No Respon'],
      'Tidak Hadir': classCounts[className]['Tidak Hadir']
    }));
  };

  const rsvpStatusData = getRsvpStatusCounts();
  const rsvpByClassData = getRsvpByClass();
  
  // Calculate totals
  const totalGuests = data.length;
  const confirmedGuests = data.filter(row => row['RSVP'] === 'Hadir').length;
  const rsvpRatio = `${confirmedGuests} / ${totalGuests}`;
  const rsvpPercentage = Math.round((confirmedGuests / totalGuests) * 100);
  
  // Colors for the pie chart
  const COLORS = ['#4CAF50', '#FFC107', '#F44336'];

  return (
    <div className="space-y-8">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-medium text-gray-700">Total Guests</h3>
          <p className="text-3xl font-bold text-blue-600 mt-2">{totalGuests}</p>
        </div>
        
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-medium text-gray-700">Confirmed Attendance</h3>
          <p className="text-3xl font-bold text-green-600 mt-2">{rsvpRatio}</p>
        </div>
        
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-medium text-gray-700">RSVP Rate</h3>
          <p className="text-3xl font-bold text-orange-600 mt-2">{rsvpPercentage}%</p>
        </div>
      </div>
      
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-medium text-gray-700 mb-4">RSVP Status Overview</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={rsvpStatusData}
                  cx="50%"
                  cy="50%"
                  labelLine={true}
                  outerRadius={80}
                  fill="#8884d8"
                  dataKey="value"
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                >
                  {rsvpStatusData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        <div className="bg-white rounded-lg shadow p-6">
          <h3 className="text-lg font-medium text-gray-700 mb-4">RSVP Status by Class (Top 10)</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart
                data={rsvpByClassData}
                margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" />
                <YAxis />
                <Tooltip />
                <Legend />
                <Bar dataKey="Hadir" fill="#4CAF50" />
                <Bar dataKey="No Respon" fill="#FFC107" />
                <Bar dataKey="Tidak Hadir" fill="#F44336" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
};

// Data Table Component
const DataTable = ({ data }) => {
  const [filteredData, setFilteredData] = useState([]);
  const [filters, setFilters] = useState({
    name: '',
    rsvp: '',
    kelas: ''
  });
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(10);

  useEffect(() => {
    filterData();
  }, [data, filters]);

  const filterData = () => {
    let filtered = [...data];
    
    // Apply name filter
    if (filters.name) {
      filtered = filtered.filter(row => 
        row['Nama'] && row['Nama'].toLowerCase().includes(filters.name.toLowerCase())
      );
    }
    
    // Apply RSVP filter
    if (filters.rsvp) {
      filtered = filtered.filter(row => {
        const status = row['RSVP'] || 'No Respon';
        return status === filters.rsvp;
      });
    }
    
    // Apply class filter
    if (filters.kelas) {
      filtered = filtered.filter(row => 
        row['Keterangan 1'] && row['Keterangan 1'].toLowerCase().includes(filters.kelas.toLowerCase())
      );
    }
    
    setFilteredData(filtered);
    setCurrentPage(1); // Reset to first page on filter change
  };

  const handleFilterChange = (e) => {
    const { name, value } = e.target;
    setFilters(prev => ({
      ...prev,
      [name]: value
    }));
  };

  // Get unique class options
  const getClassOptions = () => {
    const classes = new Set();
    data.forEach(row => {
      if (row['Keterangan 1']) {
        classes.add(row['Keterangan 1']);
      }
    });
    return Array.from(classes).sort();
  };

  // Pagination
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = filteredData.slice(indexOfFirstItem, indexOfLastItem);
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const classOptions = getClassOptions();

  return (
    <div className="bg-white rounded-lg shadow">
      <div className="p-4 border-b">
        <h2 className="text-xl font-medium text-gray-700">Guest Data</h2>
        <p className="text-gray-500 text-sm mt-1">Showing {filteredData.length} of {data.length} guests</p>
      </div>
      
      <div className="p-4 border-b bg-gray-50">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Filter by Name</label>
            <input
              type="text"
              name="name"
              value={filters.name}
              onChange={handleFilterChange}
              placeholder="Enter name..."
              className="w-full p-2 border rounded focus:ring-blue-500 focus:border-blue-500"
            />
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Filter by RSVP Status</label>
            <select
              name="rsvp"
              value={filters.rsvp}
              onChange={handleFilterChange}
              className="w-full p-2 border rounded focus:ring-blue-500 focus:border-blue-500"
            >
              <option value="">All Statuses</option>
              <option value="Hadir">Hadir</option>
              <option value="No Respon">No Respon</option>
              <option value="Tidak Hadir">Tidak Hadir</option>
            </select>
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Filter by Class</label>
            <select
              name="kelas"
              value={filters.kelas}
              onChange={handleFilterChange}
              className="w-full p-2 border rounded focus:ring-blue-500 focus:border-blue-500"
            >
              <option value="">All Classes</option>
              {classOptions.map(option => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
        </div>
      </div>
      
      <div className="overflow-x-auto">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">ID</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Class</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">RSVP Status</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">People Count</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {currentItems.length > 0 ? (
              currentItems.map((row, index) => (
                <tr key={index} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{row['Nama'] || '-'}</td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{row['ID'] || '-'}</td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{row['Keterangan 1'] || '-'}</td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${
                      row['RSVP'] === 'Hadir' 
                        ? 'bg-green-100 text-green-800' 
                        : row['RSVP'] === 'Tidak Hadir'
                          ? 'bg-red-100 text-red-800'
                          : 'bg-yellow-100 text-yellow-800'
                    }`}>
                      {row['RSVP'] || 'No Respon'}
                    </span>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{row['Jumlah Orang'] || '0'}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="5" className="px-6 py-4 text-center text-sm text-gray-500">No data found</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      
      {/* Pagination */}
      <div className="px-4 py-3 flex items-center justify-between border-t border-gray-200 sm:px-6">
        <div className="flex-1 flex justify-between sm:hidden">
          <button
            onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
            disabled={currentPage === 1}
            className={`relative inline-flex items-center px-4 py-2 border border-gray-300 text-sm font-medium rounded-md ${
              currentPage === 1 ? 'bg-gray-100 text-gray-400' : 'bg-white text-gray-700 hover:bg-gray-50'
            }`}
          >
            Previous
          </button>
          <button
            onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
            disabled={currentPage === totalPages || totalPages === 0}
            className={`ml-3 relative inline-flex items-center px-4 py-2 border border-gray-300 text-sm font-medium rounded-md ${
              currentPage === totalPages || totalPages === 0 ? 'bg-gray-100 text-gray-400' : 'bg-white text-gray-700 hover:bg-gray-50'
            }`}
          >
            Next
          </button>
        </div>
        <div className="hidden sm:flex-1 sm:flex sm:items-center sm:justify-between">
          <div>
            <p className="text-sm text-gray-700">
              Showing <span className="font-medium">{indexOfFirstItem + 1}</span> to{' '}
              <span className="font-medium">{Math.min(indexOfLastItem, filteredData.length)}</span> of{' '}
              <span className="font-medium">{filteredData.length}</span> results
            </p>
          </div>
          <div>
            <nav className="relative z-0 inline-flex rounded-md shadow-sm -space-x-px" aria-label="Pagination">
              <button
                onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                disabled={currentPage === 1}
                className={`relative inline-flex items-center px-2 py-2 rounded-l-md border border-gray-300 bg-white text-sm font-medium ${
                  currentPage === 1 ? 'text-gray-300' : 'text-gray-500 hover:bg-gray-50'
                }`}
              >
                &laquo; Previous
              </button>
              
              {/* Page numbers would go here in a more complete implementation */}
              <span className="relative inline-flex items-center px-4 py-2 border border-gray-300 bg-white text-sm font-medium">
                Page {currentPage} of {totalPages || 1}
              </span>
              
              <button
                onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                disabled={currentPage === totalPages || totalPages === 0}
                className={`relative inline-flex items-center px-2 py-2 rounded-r-md border border-gray-300 bg-white text-sm font-medium ${
                  currentPage === totalPages || totalPages === 0 ? 'text-gray-300' : 'text-gray-500 hover:bg-gray-50'
                }`}
              >
                Next &raquo;
              </button>
            </nav>
          </div>
        </div>
      </div>
    </div>
  );
};

// Testimonials Component
const Testimonials = ({ data }) => {
  const [testimonials, setTestimonials] = useState([]);
  
  useEffect(() => {
    // Filter data to only show entries with testimonials
    const filteredTestimonials = data
      .filter(row => row['Testimoni'] && row['Testimoni'].trim() !== '')
      .map(row => ({
        id: row['ID'],
        name: row['Nama'],
        kelas: row['Keterangan 1'],
        testimoni: row['Testimoni']
      }));
    
    setTestimonials(filteredTestimonials);
  }, [data]);

  return (
    <div className="space-y-8">
      <div className="bg-white rounded-lg shadow p-6">
        <h2 className="text-xl font-medium text-gray-700 mb-4">Guest Testimonials</h2>
        <p className="text-gray-500">Showing {testimonials.length} testimonials</p>
      </div>
      
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {testimonials.length > 0 ? (
          testimonials.map((item, index) => (
            <div key={index} className="bg-white rounded-lg shadow overflow-hidden">
              <div className="px-6 py-4 bg-blue-50 border-b">
                <h3 className="font-medium text-gray-800">{item.name}</h3>
                <p className="text-sm text-gray-500 mt-1">{item.kelas || 'No Class'}</p>
              </div>
              <div className="p-6">
                <p className="text-gray-700 italic">"{item.testimoni}"</p>
              </div>
            </div>
          ))
        ) : (
          <div className="col-span-3 text-center py-12 text-gray-500">
            No testimonials found
          </div>
        )}
      </div>
    </div>
  );
};

export default App;