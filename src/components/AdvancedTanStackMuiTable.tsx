import React, { useState, useEffect } from 'react';
import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  getSortedRowModel,
  getPaginationRowModel,
  getExpandedRowModel,
  useReactTable,
  ColumnDef,
  SortingState,
  ExpandedState,
  RowSelectionState,
  Row,
} from '@tanstack/react-table';
import * as XLSX from 'xlsx';
import './AdvancedMuiTable.css';

// File types and interfaces
type FileType = 'folder' | 'excel' | 'pdf' | 'doc' | 'image' | 'other';

interface FileItem {
  id: string;
  name: string;
  type: FileType;
  size: string;
  lastModified: string;
  createdBy: string;
  path: string;
  parent: string | null;
  sheetNames?: string[]; // Only for Excel files
}

// Sample data structure
const sampleFiles: FileItem[] = [
  {
    id: 'folder-1',
    name: 'Financial Reports',
    type: 'folder',
    size: '--',
    lastModified: '2025-02-15',
    createdBy: 'John Smith',
    path: '/Financial Reports',
    parent: null,
  },
  {
    id: 'folder-2',
    name: 'HR Documents',
    type: 'folder',
    size: '--',
    lastModified: '2025-02-20',
    createdBy: 'Sarah Johnson',
    path: '/HR Documents',
    parent: null,
  },
  {
    id: 'excel-1',
    name: 'Q1 Results.xlsx',
    type: 'excel',
    size: '2.4 MB',
    lastModified: '2025-03-01',
    createdBy: 'Mark Wilson',
    path: '/Q1 Results.xlsx',
    parent: null,
    sheetNames: ['Summary', 'Revenue', 'Expenses', 'Projections', 'Team Performance'],
  },
  {
    id: 'pdf-1',
    name: 'Company Policy.pdf',
    type: 'pdf',
    size: '1.8 MB',
    lastModified: '2025-01-10',
    createdBy: 'Emma Brown',
    path: '/Company Policy.pdf',
    parent: null,
  },
  {
    id: 'excel-2',
    name: 'Budget 2025.xlsx',
    type: 'excel',
    size: '3.6 MB',
    lastModified: '2025-02-28',
    createdBy: 'John Smith',
    path: '/Budget 2025.xlsx',
    parent: null,
    sheetNames: ['Overview', 'Q1', 'Q2', 'Q3', 'Q4', 'Department Breakdown', 'Historical Comparison'],
  },
];

// Nested files for folders
const nestedFiles: Record<string, FileItem[]> = {
  'folder-1': [
    {
      id: 'excel-3',
      name: 'Revenue Analysis.xlsx',
      type: 'excel',
      size: '4.2 MB',
      lastModified: '2025-02-25',
      createdBy: 'Mark Wilson',
      path: '/Financial Reports/Revenue Analysis.xlsx',
      parent: 'folder-1',
      sheetNames: ['Summary', 'Q1', 'Q2', 'Q3', 'Q4', 'Graphs'],
    },
    {
      id: 'excel-4',
      name: 'Cost Structure.xlsx',
      type: 'excel',
      size: '3.1 MB',
      lastModified: '2025-02-26',
      createdBy: 'Sarah Johnson',
      path: '/Financial Reports/Cost Structure.xlsx',
      parent: 'folder-1',
      sheetNames: ['Overview', 'Fixed Costs', 'Variable Costs', 'Comparison'],
    },
    {
      id: 'pdf-2',
      name: 'Annual Report.pdf',
      type: 'pdf',
      size: '8.7 MB',
      lastModified: '2025-02-10',
      createdBy: 'John Smith',
      path: '/Financial Reports/Annual Report.pdf',
      parent: 'folder-1',
    },
  ],
  'folder-2': [
    {
      id: 'doc-1',
      name: 'Employee Handbook.doc',
      type: 'doc',
      size: '2.8 MB',
      lastModified: '2025-01-15',
      createdBy: 'Emma Brown',
      path: '/HR Documents/Employee Handbook.doc',
      parent: 'folder-2',
    },
    {
      id: 'excel-5',
      name: 'Employee Records.xlsx',
      type: 'excel',
      size: '5.3 MB',
      lastModified: '2025-02-18',
      createdBy: 'Sarah Johnson',
      path: '/HR Documents/Employee Records.xlsx',
      parent: 'folder-2',
      sheetNames: ['All Employees', 'Engineering', 'Sales', 'Marketing', 'Management', 'HR', 'Finance'],
    },
  ],
};

const AdvancedTanStackMuiTable: React.FC = () => {
  const [sorting, setSorting] = useState<SortingState>([]);
  const [expanded, setExpanded] = useState<ExpandedState>({});
  const [rowSelection, setRowSelection] = useState<RowSelectionState>({});
  const [currentPath, setCurrentPath] = useState<string[]>([]);
  const [currentFiles, setCurrentFiles] = useState<FileItem[]>(sampleFiles);
  const [selectedSheets, setSelectedSheets] = useState<Record<string, string[]>>({});
  const [breadcrumbs, setBreadcrumbs] = useState<{id: string | null, name: string}[]>([
    { id: null, name: 'Root' }
  ]);

  // Column definitions
  const columnHelper = createColumnHelper<FileItem>();
  
  const columns = [
    // Selection column with checkbox
    columnHelper.display({
      id: 'select',
      header: ({ table }) => (
        <div className="checkbox-wrapper">
          <input
            type="checkbox"
            checked={
              table.getIsAllRowsSelected() ||
              (table.getIsSomeRowsSelected() && "indeterminate")
            }
            onChange={table.getToggleAllRowsSelectedHandler()}
            className="mui-checkbox"
            ref={(input) => {
              if (input && table.getIsSomeRowsSelected() && !table.getIsAllRowsSelected()) {
                input.indeterminate = true;
              }
            }}
          />
        </div>
      ),
      cell: ({ row }) => (
        <div className="checkbox-wrapper">
          <input
            type="checkbox"
            checked={row.getIsSelected()}
            onChange={row.getToggleSelectedHandler()}
            className="mui-checkbox"
          />
        </div>
      ),
    }),
    
    columnHelper.accessor('name', {
      header: 'Name',
      cell: ({ row, getValue }) => {
        const value = getValue();
        const fileType = row.original.type;
        
        return (
          <div className="file-cell">
            <span className={`file-icon ${fileType}-icon`}>
              <span className="material-icons">
                {fileType === 'folder' ? 'folder' : 
                 fileType === 'excel' ? 'table_chart' :
                 fileType === 'pdf' ? 'picture_as_pdf' :
                 fileType === 'doc' ? 'description' :
                 fileType === 'image' ? 'image' : 'insert_drive_file'}
              </span>
            </span>
            <span className="file-name">{value}</span>
            {fileType === 'excel' && (
              <button 
                className="mui-icon-button expand-button"
                onClick={() => {
                  row.toggleExpanded();
                }}
              >
                <span className="material-icons">
                  {row.getIsExpanded() ? 'expand_less' : 'expand_more'}
                </span>
              </button>
            )}
          </div>
        );
      },
    }),
    columnHelper.accessor('size', {
      header: 'Size',
    }),
    columnHelper.accessor('lastModified', {
      header: 'Last Modified',
      cell: info => new Date(info.getValue()).toLocaleDateString(),
    }),
    columnHelper.accessor('createdBy', {
      header: 'Created By',
    }),
    columnHelper.display({
      id: 'actions',
      header: 'Actions',
      cell: ({ row }) => (
        <div className="action-buttons">
          {row.original.type === 'folder' ? (
            <button 
              className="mui-icon-button"
              onClick={() => navigateToFolder(row.original.id)}
            >
              <span className="material-icons">open_in_new</span>
            </button>
          ) : (
            <>
              <button className="mui-icon-button">
                <span className="material-icons">download</span>
              </button>
              <button className="mui-icon-button">
                <span className="material-icons">more_vert</span>
              </button>
            </>
          )}
        </div>
      ),
    }),
  ];

  // Navigation functions
  const navigateToFolder = (folderId: string) => {
    const folder = currentFiles.find(file => file.id === folderId);
    if (folder && folder.type === 'folder') {
      const folderFiles = nestedFiles[folderId] || [];
      setCurrentFiles(folderFiles);
      setCurrentPath([...currentPath, folder.name]);
      setBreadcrumbs([...breadcrumbs, { id: folderId, name: folder.name }]);
    }
  };

  const navigateToBreadcrumb = (index: number) => {
    const newBreadcrumbs = breadcrumbs.slice(0, index + 1);
    setBreadcrumbs(newBreadcrumbs);
    
    // If navigating to root
    if (index === 0) {
      setCurrentFiles(sampleFiles);
      setCurrentPath([]);
    } else {
      const folderId = newBreadcrumbs[index].id;
      if (folderId) {
        const folderFiles = nestedFiles[folderId] || [];
        setCurrentFiles(folderFiles);
        setCurrentPath(newBreadcrumbs.slice(1).map(crumb => crumb.name));
      }
    }
  };

  // Sheet selection handler
  const handleSheetSelect = (fileId: string, sheetName: string, isSelected: boolean) => {
    setSelectedSheets(prev => {
      const currentSelection = prev[fileId] || [];
      
      if (isSelected) {
        // Add sheet to selection if not already there
        return {
          ...prev,
          [fileId]: [...currentSelection, sheetName]
        };
      } else {
        // Remove sheet from selection
        return {
          ...prev,
          [fileId]: currentSelection.filter(name => name !== sheetName)
        };
      }
    });
    
    // In a real app, you would handle the sheet selection state here
    console.log(`Sheet "${sheetName}" from file "${fileId}" selection changed to ${isSelected}`);
  };
  
  // Check if a sheet is selected
  const isSheetSelected = (fileId: string, sheetName: string): boolean => {
    return (selectedSheets[fileId] || []).includes(sheetName);
  };

  // Configure the table
  const table = useReactTable({
    data: currentFiles,
    columns,
    state: {
      sorting,
      expanded,
      rowSelection,
    },
    onSortingChange: setSorting,
    onExpandedChange: setExpanded,
    onRowSelectionChange: setRowSelection,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    getExpandedRowModel: getExpandedRowModel(),
    getRowCanExpand: row => row.original.type === 'excel',
    enableRowSelection: true,
    enableMultiRowSelection: true,
  });

  // Custom render function for expanded Excel file rows
  const renderSubComponent = ({ row }: { row: Row<FileItem> }) => {
    const file = row.original;
    if (file.type !== 'excel' || !file.sheetNames) return null;

    return (
      <div className="excel-sheets-container">
        <div className="excel-sheets-header">
          <span className="material-icons">article</span>
          <span>Sheets in {file.name}</span>
          <button 
            className="mui-button mui-button-secondary sheet-select-all"
            onClick={() => {
              // Toggle all sheets selection
              const allSelected = file.sheetNames?.every(sheet => isSheetSelected(file.id, sheet));
              
              if (allSelected) {
                // If all are selected, clear selection
                setSelectedSheets(prev => ({
                  ...prev,
                  [file.id]: []
                }));
              } else {
                // If not all selected, select all
                setSelectedSheets(prev => ({
                  ...prev,
                  [file.id]: [...(file.sheetNames || [])]
                }));
              }
            }}
          >
            {file.sheetNames?.every(sheet => isSheetSelected(file.id, sheet))
              ? 'Deselect All'
              : 'Select All'}
          </button>
        </div>
        <div className="excel-sheets-list">
          {file.sheetNames.map(sheet => (
            <div 
              key={sheet}
              className={`excel-sheet-item ${isSheetSelected(file.id, sheet) ? 'selected' : ''}`}
            >
              <input
                type="checkbox"
                checked={isSheetSelected(file.id, sheet)}
                onChange={e => handleSheetSelect(file.id, sheet, e.target.checked)}
                className="mui-checkbox sheet-checkbox"
                onClick={e => e.stopPropagation()}
              />
              <span className="material-icons sheet-icon">table_chart</span>
              <span 
                className="sheet-name"
                onClick={() => handleSheetSelect(file.id, sheet, !isSheetSelected(file.id, sheet))}
              >
                {sheet}
              </span>
            </div>
          ))}
        </div>
      </div>
    );
  };

  return (
    <div className="mui-table-container">
      {/* Directory navigation */}
      <div className="mui-directory-nav">
        <div className="mui-breadcrumbs">
          {breadcrumbs.map((crumb, index) => (
            <React.Fragment key={index}>
              {index > 0 && <span className="breadcrumb-separator">/</span>}
              <span 
                className="breadcrumb-item"
                onClick={() => navigateToBreadcrumb(index)}
              >
                {crumb.name}
              </span>
            </React.Fragment>
          ))}
        </div>
      </div>
      
      {/* Toolbar */}
      <div className="mui-table-toolbar">
        <div className="mui-toolbar-left">
          <h2>Files</h2>
          {Object.keys(rowSelection).length > 0 && (
            <div className="selected-count">
              {Object.keys(rowSelection).length} items selected
            </div>
          )}
        </div>
        <div className="mui-table-actions">
          {Object.keys(rowSelection).length > 0 ? (
            <>
              <button className="mui-button mui-button-secondary">
                <span className="material-icons">download</span>
                Download
              </button>
              <button className="mui-button mui-button-secondary">
                <span className="material-icons">delete</span>
                Delete
              </button>
            </>
          ) : (
            <>
              <div className="mui-search">
                <span className="material-icons">search</span>
                <input type="text" placeholder="Search files..." />
              </div>
              <button className="mui-button mui-button-primary">
                <span className="material-icons">upload_file</span>
                Upload
              </button>
              <button className="mui-button mui-button-secondary">
                <span className="material-icons">create_new_folder</span>
                New Folder
              </button>
            </>
          )}
        </div>
      </div>
      
      {/* Table */}
      <div className="mui-table-wrapper">
        <table className="mui-table">
          <thead>
            {table.getHeaderGroups().map(headerGroup => (
              <tr key={headerGroup.id}>
                {headerGroup.headers.map(header => (
                  <th key={header.id} className={header.column.getCanSort() ? 'sortable' : ''}>
                    <div 
                      className="th-content"
                      onClick={header.column.getToggleSortingHandler()}
                    >
                      {flexRender(
                        header.column.columnDef.header,
                        header.getContext()
                      )}
                      {header.column.getCanSort() && (
                        <span className="sort-icon">
                          {header.column.getIsSorted() === 'asc' ? (
                            <span className="material-icons">arrow_upward</span>
                          ) : header.column.getIsSorted() === 'desc' ? (
                            <span className="material-icons">arrow_downward</span>
                          ) : (
                            <span className="material-icons sort-inactive">unfold_more</span>
                          )}
                        </span>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.map(row => (
              <React.Fragment key={row.id}>
                <tr 
                  className={`mui-table-row ${row.original.type === 'folder' ? 'folder-row' : ''}`}
                  onDoubleClick={() => {
                    if (row.original.type === 'folder') {
                      navigateToFolder(row.original.id);
                    }
                  }}
                >
                  {row.getVisibleCells().map(cell => (
                    <td key={cell.id}>
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
                {/* Expanded content for Excel files */}
                {row.getIsExpanded() && (
                  <tr className="expanded-row">
                    <td colSpan={columns.length}>
                      {renderSubComponent({ row })}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            ))}
            {currentFiles.length === 0 && (
              <tr>
                <td colSpan={columns.length} className="empty-folder">
                  <div className="empty-folder-content">
                    <span className="material-icons">folder_open</span>
                    <span>This folder is empty</span>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div className="mui-table-pagination">
        <div className="mui-rows-per-page">
          <span>Rows per page:</span>
          <select
            value={table.getState().pagination.pageSize}
            onChange={e => {
              table.setPageSize(Number(e.target.value));
            }}
          >
            {[5, 10, 20, 50].map(pageSize => (
              <option key={pageSize} value={pageSize}>
                {pageSize}
              </option>
            ))}
          </select>
        </div>
        
        <div className="mui-pagination-info">
          {table.getState().pagination.pageIndex * table.getState().pagination.pageSize + 1}-
          {Math.min(
            (table.getState().pagination.pageIndex + 1) * table.getState().pagination.pageSize,
            table.getPrePaginationRowModel().rows.length
          )}{' '}
          of {table.getPrePaginationRowModel().rows.length}
        </div>
        
        <div className="mui-pagination-actions">
          <button
            className="mui-icon-button"
            onClick={() => table.setPageIndex(0)}
            disabled={!table.getCanPreviousPage()}
          >
            <span className="material-icons">first_page</span>
          </button>
          <button
            className="mui-icon-button"
            onClick={() => table.previousPage()}
            disabled={!table.getCanPreviousPage()}
          >
            <span className="material-icons">chevron_left</span>
          </button>
          <button
            className="mui-icon-button"
            onClick={() => table.nextPage()}
            disabled={!table.getCanNextPage()}
          >
            <span className="material-icons">chevron_right</span>
          </button>
          <button
            className="mui-icon-button"
            onClick={() => table.setPageIndex(table.getPageCount() - 1)}
            disabled={!table.getCanNextPage()}
          >
            <span className="material-icons">last_page</span>
          </button>
        </div>
      </div>
    </div>
  );
};

export default AdvancedTanStackMuiTable;