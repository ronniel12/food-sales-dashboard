import React, { useState, useEffect } from 'react';
import { 
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, 
  Tooltip, Legend, ResponsiveContainer, ComposedChart
} from 'recharts';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import _ from 'lodash';

const FoodSalesDashboard = () => {
  const [loading, setLoading] = useState(true);
  const [dishes, setDishes] = useState([]);
  const [dishData, setDishData] = useState([]);
  const [selectedDish, setSelectedDish] = useState('');
  const [months, setMonths] = useState([]);
  const [viewMode, setViewMode] = useState('individual');
  const [topDishes, setTopDishes] = useState([]);
  const [exportingData, setExportingData] = useState(false);

  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        const response = await fetch('/Food Monitoring scale.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        const workbook = XLSX.read(arrayBuffer, {
          cellStyles: true,
          cellFormulas: true,
          cellDates: true,
          cellNF: true,
          sheetStubs: true
        });

        // Get data from the DISH sheet
        const dishSheet = workbook.Sheets["DISH"];
        const dishJsonData = XLSX.utils.sheet_to_json(dishSheet, { defval: null });
        
        // Extract months
        const monthColumns = Object.keys(dishJsonData[0])
          .filter(key => key !== "Dish " && key !== "__EMPTY");
        
        setMonths(monthColumns);
        
        // Prepare dish data
        const allDishes = dishJsonData.map(row => row["Dish "]);
        setDishes(allDishes);
        setSelectedDish(allDishes[0]);  // Set first dish as default
        
        // Transform data for charts
        const transformedData = transformDataForCharts(dishJsonData, monthColumns);
        setDishData(transformedData);
        
        // Identify top 5 dishes
        const totalSalesByDish = {};
        dishJsonData.forEach(row => {
          const dishName = row["Dish "];
          const totalSales = monthColumns.reduce((sum, month) => sum + (row[month] || 0), 0);
          totalSalesByDish[dishName] = totalSales;
        });
        
        const sortedDishes = Object.entries(totalSalesByDish)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 5)
          .map(entry => entry[0]);
          
        setTopDishes(sortedDishes);
        
        setLoading(false);
      } catch (error) {
        console.error("Error loading data:", error);
        setLoading(false);
      }
    };
    
    loadData();
  }, []);
  
  // Transform data for chart visualization
  const transformDataForCharts = (rawData, months) => {
    const transformedData = {};
    
    rawData.forEach(row => {
      const dishName = row["Dish "];
      const dishMonthlyData = months.map(month => ({
        month,
        sales: row[month] || 0
      }));
      
      transformedData[dishName] = dishMonthlyData;
    });
    
    return transformedData;
  };
  
  // Calculate growth rates for a dish
  const calculateGrowthRates = (dishName) => {
    if (!dishData[dishName]) return [];
    
    const data = dishData[dishName];
    const growthData = [];
    
    for (let i = 1; i < data.length; i++) {
      const prevSales = data[i-1].sales;
      const currentSales = data[i].sales;
      
      let growthRate = 0;
      if (prevSales > 0) {
        growthRate = ((currentSales - prevSales) / prevSales) * 100;
      }
      
      growthData.push({
        month: data[i].month,
        growth: growthRate.toFixed(1)
      });
    }
    
    return growthData;
  };
  
  // Get color based on growth rate
  const getGrowthColor = (growth) => {
    if (growth > 0) return "#4CAF50";  // Green for positive growth
    if (growth < 0) return "#F44336";  // Red for negative growth
    return "#9E9E9E";  // Gray for no change
  };
  
  // Export data to Excel
  const exportToExcel = async () => {
    setExportingData(true);
    console.log('Starting Excel export...');
    
    try {
      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Food Sales Dashboard';
      workbook.lastModifiedBy = 'Food Sales Dashboard';
      workbook.created = new Date();
      workbook.modified = new Date();

      console.log('Workbook created');

      // Create overview sheet with all dishes
      const overviewSheet = workbook.addWorksheet('All Dishes Overview');
      console.log('Overview sheet created');
      
      // Add headers
      const headers = ['Dish', ...months];
      overviewSheet.addRow(headers);
      
      // Add data
      dishes.forEach(dish => {
        const row = [dish];
        months.forEach(month => {
          const dishMonthData = dishData[dish]?.find(m => m.month === month);
          row.push(dishMonthData?.sales || 0);
        });
        overviewSheet.addRow(row);
      });

      // Format the overview sheet
      overviewSheet.getRow(1).font = { bold: true };
      overviewSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
      };

      // Create individual sheets for top dishes
      topDishes.forEach(dish => {
        const growthData = calculateGrowthRates(dish);
        const individualSheet = workbook.addWorksheet(dish.substring(0, 30));
        
        // Add headers
        individualSheet.addRow(['Month', 'Sales', 'Growth Rate (%)']);
        
        // Add data
        months.forEach((month, index) => {
          const sales = dishData[dish]?.find(m => m.month === month)?.sales || 0;
          const growth = index > 0 ? growthData.find(g => g.month === month)?.growth || 0 : 'N/A';
          individualSheet.addRow([month, sales, growth]);
        });

        // Format the individual sheet
        individualSheet.getRow(1).font = { bold: true };
        individualSheet.getRow(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' }
        };
      });

      // Create summary analysis sheet
      const summarySheet = workbook.addWorksheet('Sales Summary');
      
      // Add headers
      summarySheet.addRow(['Dish', 'Total Sales', 'Average Monthly Sales', 'Overall Trend (%)', 'First Month Sales', 'Last Month Sales']);
      
      // Add data
      const summaryData = dishes.map(dish => {
        const monthlySales = months.map(month => {
          return dishData[dish]?.find(m => m.month === month)?.sales || 0;
        });
        
        const totalSales = monthlySales.reduce((sum, sales) => sum + sales, 0);
        const avgSales = totalSales / months.length;
        
        const firstMonth = monthlySales[0];
        const lastMonth = monthlySales[monthlySales.length - 1];
        let trend = 0;
        
        if (firstMonth > 0) {
          trend = ((lastMonth - firstMonth) / firstMonth) * 100;
        }
        
        return {
          dish,
          totalSales,
          avgSales,
          trend,
          firstMonth,
          lastMonth
        };
      }).sort((a, b) => b.totalSales - a.totalSales);

      summaryData.forEach(row => {
        summarySheet.addRow([
          row.dish,
          row.totalSales,
          row.avgSales.toFixed(1),
          row.trend.toFixed(1),
          row.firstMonth,
          row.lastMonth
        ]);
      });

      // Format the summary sheet
      summarySheet.getRow(1).font = { bold: true };
      summarySheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
      };

      // Add conditional formatting for trend values
      summaryData.forEach((row, index) => {
        const trendCell = summarySheet.getCell(`D${index + 2}`);
        trendCell.font = {
          color: { argb: row.trend >= 0 ? 'FF008000' : 'FFFF0000' }
        };
      });

      // Write to file and trigger download
      console.log('Writing workbook to buffer...');
      const buffer = await workbook.xlsx.writeBuffer();
      console.log('Buffer created, creating blob...');
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      console.log('Blob URL created:', url);
      
      const link = document.createElement('a');
      link.href = url;
      link.download = 'Food_Sales_Analysis.xlsx';
      console.log('Triggering download...');
      link.click();
      window.URL.revokeObjectURL(url);
      console.log('Download triggered');
    } catch (error) {
      console.error("Error exporting to Excel:", error);
      alert("Error exporting to Excel: " + error.message);
    } finally {
      setExportingData(false);
    }
  };
  
  return (
    <div className="p-6 max-w-full mx-auto bg-gray-50 rounded-lg shadow-md">
      <h1 className="text-2xl font-bold mb-6 text-center">Food Sales Analysis Dashboard</h1>
      
      {loading ? (
        <div className="text-center py-8">
          <p className="text-lg">Loading data...</p>
        </div>
      ) : (
        <div>
          <div className="flex flex-col md:flex-row justify-between items-center mb-6">
            <div className="w-full md:w-1/3 mb-4 md:mb-0">
              <div className="flex space-x-4">
                <button 
                  onClick={() => setViewMode('individual')}
                  className={`px-4 py-2 rounded ${viewMode === 'individual' ? 'bg-blue-600 text-white' : 'bg-gray-200'}`}
                >
                  Individual Dish
                </button>
                <button 
                  onClick={() => setViewMode('comparison')}
                  className={`px-4 py-2 rounded ${viewMode === 'comparison' ? 'bg-blue-600 text-white' : 'bg-gray-200'}`}
                >
                  Comparison
                </button>
              </div>
            </div>
            
            {viewMode === 'individual' && (
              <div className="w-full md:w-1/3">
                <select 
                  value={selectedDish} 
                  onChange={(e) => setSelectedDish(e.target.value)}
                  className="w-full p-2 border rounded bg-white"
                >
                  {dishes.map(dish => (
                    <option key={dish} value={dish}>{dish}</option>
                  ))}
                </select>
              </div>
            )}
            
            <div className="w-full md:w-1/3 text-right">
              <button 
                onClick={exportToExcel}
                disabled={exportingData}
                className="px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700 disabled:bg-gray-400"
              >
                {exportingData ? 'Exporting...' : 'Export to Excel'}
              </button>
            </div>
          </div>
          
          {viewMode === 'individual' && selectedDish && (
            <div>
              <div className="bg-white p-4 rounded-lg shadow mb-6">
                <h2 className="text-xl font-semibold mb-4">{selectedDish} - Monthly Sales</h2>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart 
                      data={dishData[selectedDish]} 
                      margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="month" />
                      <YAxis />
                      <Tooltip />
                      <Legend />
                      <Line 
                        type="monotone" 
                        dataKey="sales" 
                        stroke="#3B82F6" 
                        activeDot={{ r: 8 }} 
                        name="Sales"
                      />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
              
              <div className="bg-white p-4 rounded-lg shadow mb-6">
                <h2 className="text-xl font-semibold mb-4">{selectedDish} - Growth Rate</h2>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={calculateGrowthRates(selectedDish)} 
                      margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="month" />
                      <YAxis unit="%" />
                      <Tooltip formatter={(value) => [`${value}%`, 'Growth Rate']} />
                      <Legend />
                      <Bar 
                        dataKey="growth" 
                        name="Growth Rate" 
                        fill="#3B82F6"
                        radius={[4, 4, 0, 0]}
                      >
                        {calculateGrowthRates(selectedDish).map((entry, index) => (
                          <div key={`cell-${index}`} fill={getGrowthColor(parseFloat(entry.growth))} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              
              {/* Monthly Sales Data Table */}
              <div className="bg-white p-4 rounded-lg shadow">
                <h2 className="text-xl font-semibold mb-4">{selectedDish} - Sales Data</h2>
                <div className="overflow-x-auto">
                  <table className="min-w-full bg-white">
                    <thead className="bg-gray-100">
                      <tr>
                        <th className="py-2 px-4 border-b text-left">Month</th>
                        <th className="py-2 px-4 border-b text-right">Sales</th>
                        <th className="py-2 px-4 border-b text-right">Growth Rate</th>
                      </tr>
                    </thead>
                    <tbody>
                      {months.map((month, index) => {
                        const data = dishData[selectedDish]?.find(m => m.month === month);
                        const sales = data ? data.sales : 0;
                        
                        // Calculate growth rate
                        let growthRate = 'N/A';
                        let growthClass = '';
                        
                        if (index > 0) {
                          const prevMonthData = dishData[selectedDish]?.find(m => m.month === months[index - 1]);
                          const prevSales = prevMonthData ? prevMonthData.sales : 0;
                          
                          if (prevSales > 0) {
                            const growth = ((sales - prevSales) / prevSales) * 100;
                            growthRate = `${growth.toFixed(1)}%`;
                            
                            if (growth > 0) growthClass = 'text-green-600';
                            else if (growth < 0) growthClass = 'text-red-600';
                          }
                        }
                        
                        return (
                          <tr key={month} className="hover:bg-gray-50">
                            <td className="py-2 px-4 border-b">{month}</td>
                            <td className="py-2 px-4 border-b text-right">{sales}</td>
                            <td className={`py-2 px-4 border-b text-right ${growthClass}`}>{growthRate}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
          
          {viewMode === 'comparison' && (
            <div>
              <div className="bg-white p-4 rounded-lg shadow mb-6">
                <h2 className="text-xl font-semibold mb-4">Top 5 Dishes Comparison</h2>
                <div className="h-96">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis 
                        dataKey="month" 
                        type="category" 
                        allowDuplicatedCategory={false} 
                        data={months.map(month => ({ month }))}
                      />
                      <YAxis />
                      <Tooltip />
                      <Legend />
                      
                      {topDishes.map((dish, index) => (
                        <Line 
                          key={dish}
                          data={dishData[dish]}
                          type="monotone"
                          dataKey="sales"
                          name={dish}
                          stroke={[
                            "#3B82F6", "#10B981", "#EF4444", 
                            "#F59E0B", "#8B5CF6"
                          ][index % 5]}
                          activeDot={{ r: 8 }}
                          strokeWidth={2}
                        />
                      ))}
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
              
              <div className="bg-white p-4 rounded-lg shadow mb-6">
                <h2 className="text-xl font-semibold mb-4">Total Sales by Dish</h2>
                <div className="h-96">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      layout="vertical"
                      data={topDishes.map(dish => {
                        const totalSales = (dishData[dish] || [])
                          .reduce((sum, item) => sum + item.sales, 0);
                        return { dish, totalSales };
                      })}
                      margin={{ top: 5, right: 30, left: 100, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis type="number" />
                      <YAxis dataKey="dish" type="category" />
                      <Tooltip />
                      <Legend />
                      <Bar 
                        dataKey="totalSales" 
                        fill="#3B82F6" 
                        name="Total Sales"
                        radius={[0, 4, 4, 0]}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              
              {/* Comparison Table */}
              <div className="bg-white p-4 rounded-lg shadow">
                <h2 className="text-xl font-semibold mb-4">Sales Comparison Table</h2>
                <div className="overflow-x-auto">
                  <table className="min-w-full bg-white">
                    <thead className="bg-gray-100">
                      <tr>
                        <th className="py-2 px-4 border-b text-left">Dish</th>
                        {months.map(month => (
                          <th key={month} className="py-2 px-4 border-b text-right">{month}</th>
                        ))}
                        <th className="py-2 px-4 border-b text-right">Total</th>
                        <th className="py-2 px-4 border-b text-right">Trend</th>
                      </tr>
                    </thead>
                    <tbody>
                      {topDishes.map(dish => {
                        const monthlySales = months.map(month => {
                          return dishData[dish]?.find(m => m.month === month)?.sales || 0;
                        });
                        
                        const totalSales = monthlySales.reduce((sum, sales) => sum + sales, 0);
                        
                        // Calculate trend (comparing first and last month)
                        const firstMonth = monthlySales[0];
                        const lastMonth = monthlySales[monthlySales.length - 1];
                        let trend = 0;
                        let trendClass = '';
                        
                        if (firstMonth > 0) {
                          trend = ((lastMonth - firstMonth) / firstMonth) * 100;
                          if (trend > 0) trendClass = 'text-green-600';
                          else if (trend < 0) trendClass = 'text-red-600';
                        }
                        
                        return (
                          <tr key={dish} className="hover:bg-gray-50">
                            <td className="py-2 px-4 border-b font-medium">{dish}</td>
                            
                            {months.map(month => {
                              const sales = dishData[dish]?.find(m => m.month === month)?.sales || 0;
                              return (
                                <td key={`${dish}-${month}`} className="py-2 px-4 border-b text-right">
                                  {sales}
                                </td>
                              );
                            })}
                            
                            <td className="py-2 px-4 border-b text-right font-semibold">{totalSales}</td>
                            <td className={`py-2 px-4 border-b text-right font-semibold ${trendClass}`}>
                              {trend.toFixed(1)}%
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default FoodSalesDashboard; 