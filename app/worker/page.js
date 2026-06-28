'use client';

import { useState, useEffect } from 'react';
import { supabase } from '../../utils/supabase';

export default function WorkerInterface() {
  // --- STATE MANAGEMENT ---
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(true);
  const [updatingId, setUpdatingId] = useState(null);
  const [currentFilter, setCurrentFilter] = useState('All');

  // --- FETCH TASKS ENGINE ---
  async function fetchTasks() {
    setLoading(true);
    const { data, error } = await supabase
      .from('products')
      .select('*')
      .order('updated_at', { ascending: false });

    if (error) {
      console.error('Error fetching catalog tasks:', error);
    } else {
      setTasks(data || []);
    }
    setLoading(false);
  }

  useEffect(() => {
    fetchTasks();
  }, []);

  // --- PIPELINE STATE MACHINE UPDATE ---
  const handleStatusUpdate = async (id, currentStatus) => {
    let nextStatus = 'Missing';
    
    // Simple pipeline progression: Missing -> Processing -> Completed
    if (currentStatus === 'Missing') nextStatus = 'Processing';
    if (currentStatus === 'Processing') nextStatus = 'Completed';
    if (currentStatus === 'Completed') nextStatus = 'Missing'; // Allows cycling back if needed

    setUpdatingId(id);

    const { error } = await supabase
      .from('products')
      .update({ status: nextStatus })
      .eq('id', id);

    if (error) {
      alert('Failed to update task progress: ' + error.message);
    } else {
      // Optimistically update local UI state immediately for responsive feedback
      setTasks(prevTasks =>
        prevTasks.map(task =>
          task.id === id ? { ...task, status: nextStatus } : task
        )
      );
    }
    setUpdatingId(null);
  };

  // --- FILTER CHIP LOGIC ---
  const filteredTasks = tasks.filter(task => {
    if (currentFilter === 'All') return true;
    return task.status.toLowerCase() === currentFilter.toLowerCase();
  });

  // Count helper for UI statistic cards
  const getCount = (status) => tasks.filter(t => t.status === status).length;

  return (
    <div className="min-h-screen bg-gray-100 p-4 sm:p-8 font-sans">
      <div className="max-w-6xl mx-auto">
        
        {/* Portal Header */}
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between border-b border-gray-200 pb-5 mb-6 gap-4">
          <div>
            <h1 className="text-2xl sm:text-3xl font-black text-gray-900 tracking-tight uppercase">
              ⚙️ Task Execution Center
            </h1>
            <p className="text-sm text-gray-500 font-medium mt-0.5">
              Black-rose Fulfillment & Catalog Operations Portal
            </p>
          </div>
          
          {/* Live System Statistics Dashboard */}
          <div className="flex gap-2 text-center">
            <div className="bg-white px-3 py-2 rounded-lg border border-gray-200 shadow-xs">
              <div className="text-xs font-bold uppercase text-gray-400">Missing</div>
              <div className="text-lg font-black text-amber-600">{getCount('Missing')}</div>
            </div>
            <div className="bg-white px-3 py-2 rounded-lg border border-gray-200 shadow-xs">
              <div className="text-xs font-bold uppercase text-gray-400">Processing</div>
              <div className="text-lg font-black text-blue-600">{getCount('Processing')}</div>
            </div>
            <div className="bg-white px-3 py-2 rounded-lg border border-gray-200 shadow-xs">
              <div className="text-xs font-bold uppercase text-gray-400">Done</div>
              <div className="text-lg font-black text-green-600">{getCount('Completed')}</div>
            </div>
          </div>
        </div>

        {/* Filter Navigation Control Bar */}
        <div className="flex flex-wrap gap-2 mb-6">
          {['All', 'Missing', 'Processing', 'Completed'].map((filter) => (
            <button
              key={filter}
              onClick={() => setCurrentFilter(filter)}
              className={`px-4 py-2 text-xs font-bold uppercase tracking-wider rounded-lg border transition-all ${
                currentFilter === filter
                  ? 'bg-gray-900 text-white border-gray-900 shadow-sm'
                  : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-50'
              }`}
            >
              {filter} {filter === 'All' ? `(${tasks.length})` : `(${getCount(filter)})`}
            </button>
          ))}
        </div>

        {/* Active Task Roster Matrix */}
        {loading ? (
          <div className="text-center py-16 text-gray-400 bg-white rounded-xl border border-gray-200 shadow-xs font-medium">
            Syncing catalog queues with cloud tables...
          </div>
        ) : filteredTasks.length === 0 ? (
          <div className="text-center py-20 text-gray-400 bg-white rounded-xl border border-gray-200 shadow-xs">
            <div className="text-4xl mb-2">🎉</div>
            <p className="font-bold text-gray-700">No active tasks found in this queue.</p>
            <p className="text-xs text-gray-400 mt-1">Excellent work keeping the pipeline clear!</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {filteredTasks.map((task) => (
              <div 
                key={task.id} 
                className={`bg-white rounded-xl border p-5 shadow-xs transition-all flex flex-col justify-between ${
                  task.status === 'Completed' ? 'border-green-200 bg-green-50/10' : 'border-gray-200'
                }`}
              >
                <div>
                  {/* Card Header Info */}
                  <div className="flex justify-between items-start gap-2 mb-3">
                    <span className="font-mono text-xs font-bold tracking-wider bg-gray-100 text-gray-600 px-2 py-1 rounded">
                      {task.sku}
                    </span>
                    <span className={`text-xxs uppercase tracking-widest font-black px-2 py-1 rounded-sm border ${
                      task.status === 'Missing' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                      task.status === 'Processing' ? 'bg-blue-50 text-blue-700 border-blue-200' :
                      'bg-green-50 text-green-700 border-green-200'
                    }`}>
                      {task.status}
                    </span>
                  </div>

                  {/* Product Details */}
                  <h3 className="text-lg font-bold text-gray-900 leading-tight mb-1">
                    {task.product_name}
                  </h3>
                  <div className="text-xs font-semibold text-gray-400 mb-4 uppercase tracking-wider">
                    {task.category} &bull; <span className="text-gray-500">{task.warehouse}</span>
                  </div>

                  {/* Target Outlets Display */}
                  <div className="mb-6">
                    <span className="text-xxs font-bold text-gray-400 uppercase tracking-wider block mb-1.5">
                      Fulfillment Channels
                    </span>
                    <div className="flex flex-wrap gap-1">
                      {task.platforms?.map((plat, idx) => (
                        <span key={idx} className="bg-gray-50 border border-gray-200 text-gray-600 text-xxs font-bold px-2 py-0.5 rounded">
                          {plat}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Real-time State Control Action Button */}
                <button
                  disabled={updatingId === task.id}
                  onClick={() => handleStatusUpdate(task.id, task.status)}
                  className={`w-full py-3 px-4 rounded-xl text-xs font-black uppercase tracking-wider transition-all border shadow-xs flex items-center justify-center gap-2 ${
                    updatingId === task.id ? 'bg-gray-50 text-gray-400 border-gray-200 cursor-wait' :
                    task.status === 'Missing' ? 'bg-amber-600 hover:bg-amber-700 text-white border-amber-600' :
                    task.status === 'Processing' ? 'bg-blue-600 hover:bg-blue-700 text-white border-blue-600' :
                    'bg-white text-gray-700 hover:bg-gray-50 border-gray-300'
                  }`}
                >
                  {updatingId === task.id ? (
                    'Writing to cloud...'
                  ) : task.status === 'Missing' ? (
                    <>🚀 Start Processing</>
                  ) : task.status === 'Processing' ? (
                    <>✅ Mark As Completed</>
                  ) : (
                    <>🔄 Re-open Task</>
                  )}
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}