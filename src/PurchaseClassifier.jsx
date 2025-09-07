import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Undo2, Download, ChevronLeft, ChevronRight, FileSpreadsheet, AlertCircle, Check, Database, History, Trash2, Search, Calendar, TrendingUp, SkipForward, Split, Edit2, DollarSign, Building2 } from 'lucide-react';

const PurchaseClassifier = () => {
  const [file, setFile] = useState(null);
  const [rawData, setRawData] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [classifications, setClassifications] = useState({});
  const [skipped, setSkipped] = useState(new Set());
  const [splitTransactions, setSplitTransactions] = useState({});
  const [history, setHistory] = useState([]);
  const [columns, setColumns] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [headerRow, setHeaderRow] = useState(0);
  const [dataStartRow, setDataStartRow] = useState(1);
  const [fileMonth, setFileMonth] = useState('');
  const [fileYear, setFileYear] = useState(new Date().getFullYear());
  const [suggestion, setSuggestion] = useState(null);
  const [showDatabase, setShowDatabase] = useState(false);
  const [historicalData, setHistoricalData] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [showSplitModal, setShowSplitModal] = useState(false);
  const [splitRatio, setSplitRatio] = useState(50);
  const [splitClasses, setSplitClasses] = useState({ first: '', second: '' });
  const [isReclassifying, setIsReclassifying] = useState(false);
  const [dateColumn, setDateColumn] = useState(null);
  const [amountColumn, setAmountColumn] = useState(null);
  const [showClassSelection, setShowClassSelection] = useState(true);
  const [companyLogos, setCompanyLogos] = useState({});
  const [autoDetectedMonth, setAutoDetectedMonth] = useState(false);

  // helper to parse amounts with comma or dot decimal
  const parseAmount = (value) => {
    if (typeof value === 'number') return value;
    if (!value) return 0;
    const normalized = String(value)
      .replace(/\s+/g, '')
      .replace(/,/, '.')
      .replace(/[^0-9.-]/g, '');
    const num = parseFloat(normalized);
    return isNaN(num) ? 0 : num;
  };

  // Load historical data on component mount
  useEffect(() => {
    loadHistoricalData();
  }, []);

  // Update suggestion when transaction changes
  const currentTransaction = transactions[currentIndex];
  useEffect(() => {
    if (currentTransaction && !isReclassifying) {
      const suggested = getSuggestion(currentTransaction);
      setSuggestion(suggested);
      fetchCompanyLogo(currentTransaction);
    } else {
      setSuggestion(null);
    }
  }, [currentTransaction, isReclassifying]);

  const fetchCompanyLogo = async (transaction) => {
    const companyFields = ['Mottagare', 'Vendor', 'Merchant', 'Store', 'Supplier', 'Name', 'Beskrivning'];
    let companyName = null;
    for (const field of companyFields) {
      if (transaction[field] && transaction[field].length > 2) {
        companyName = transaction[field];
        break;
      }
    }
    if (!companyName || companyLogos[companyName]) return;
    const cleanName = companyName.toLowerCase().replace(/[^a-z0-9\s]/g, '').split(' ')[0];
    const logoUrl = `https://logo.clearbit.com/${cleanName}.com`;
    try {
      const img = new Image();
      img.onload = () => {
        setCompanyLogos(prev => ({ ...prev, [companyName]: logoUrl }));
      };
      img.onerror = () => {
        const seLogoUrl = `https://logo.clearbit.com/${cleanName}.se`;
        const img2 = new Image();
        img2.onload = () => {
          setCompanyLogos(prev => ({ ...prev, [companyName]: seLogoUrl }));
        };
        img2.src = seLogoUrl;
      };
      img.src = logoUrl;
    } catch (e) {
      console.log('Could not fetch logo for', companyName);
    }
  };

  const loadHistoricalData = () => {
    const stored = localStorage.getItem('purchaseClassifications');
    if (stored) {
      try {
        const data = JSON.parse(stored);
        setHistoricalData(data);
      } catch (e) {
        console.error('Error loading historical data:', e);
        setHistoricalData([]);
      }
    }
  };

  const saveToDatabase = (classifiedTransactions) => {
    const stored = localStorage.getItem('purchaseClassifications');
    let existingData = [];
    if (stored) {
      try {
        existingData = JSON.parse(stored);
      } catch (e) {
        console.error('Error parsing stored data:', e);
      }
    }
    const dataToSave = classifiedTransactions.map(transaction => ({
      ...transaction,
      importMonth: fileMonth,
      importYear: fileYear,
      importDate: new Date().toISOString()
    }));
    const updatedData = [...existingData, ...dataToSave];
    localStorage.setItem('purchaseClassifications', JSON.stringify(updatedData));
    setHistoricalData(updatedData);
    return updatedData.length;
  };

  const clearDatabase = () => {
    if (window.confirm('Are you sure you want to clear all historical data? This cannot be undone.')) {
      localStorage.removeItem('purchaseClassifications');
      setHistoricalData([]);
    }
  };

  const getSuggestion = (transaction) => {
    const stored = localStorage.getItem('purchaseClassifications');
    if (!stored) return null;
    try {
      const historicalData = JSON.parse(stored);
      const matchFields = ['Description', 'Vendor', 'Merchant', 'Store', 'Supplier', 'Name', 'Title', 'Beskrivning', 'Mottagare'];
      for (const field of matchFields) {
        if (transaction[field]) {
          const transactionValue = String(transaction[field]).toLowerCase().trim();
          const matches = historicalData.filter(item => {
            const itemValue = String(item[field] || '').toLowerCase().trim();
            return itemValue && (itemValue === transactionValue || itemValue.includes(transactionValue) || transactionValue.includes(itemValue));
          });
          if (matches.length > 0) {
            const classCount = {};
            matches.forEach(match => {
              if (match.Classification) {
                classCount[match.Classification] = (classCount[match.Classification] || 0) + 1;
              }
            });
            const sorted = Object.entries(classCount).sort((a, b) => b[1] - a[1]);
            if (sorted.length > 0) {
              return {
                classification: sorted[0][0],
                confidence: (sorted[0][1] / matches.length * 100).toFixed(0),
                matchCount: matches.length,
                field: field
              };
            }
          }
        }
      }
    } catch (e) {
      console.error('Error getting suggestion:', e);
    }
    return null;
  };

  const excelDateToJS = (excelDate) => {
    if (typeof excelDate === 'number') {
      const daysSince1900 = excelDate;
      const millisecondsPerDay = 24 * 60 * 60 * 1000;
      const adjustedDays = daysSince1900 > 60 ? daysSince1900 - 2 : daysSince1900 - 1;
      const baseDate = new Date(1900, 0, 1);
      const resultDate = new Date(baseDate.getTime() + adjustedDays * millisecondsPerDay);
      const year = resultDate.getFullYear();
      const month = String(resultDate.getMonth() + 1).padStart(2, '0');
      const day = String(resultDate.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    if (typeof excelDate === 'string') {
      if (/^\d{4}-\d{2}-\d{2}$/.test(excelDate)) {
        return excelDate;
      }
      const parsed = new Date(excelDate);
      if (!isNaN(parsed)) {
        const year = parsed.getFullYear();
        const month = String(parsed.getMonth() + 1).padStart(2, '0');
        const day = String(parsed.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }
    return excelDate;
  };

  const extractMonthYear = (rawJsonData, headerRowIndex) => {
    const headers = rawJsonData[headerRowIndex];
    let valutaColumnIndex = -1;
    headers.forEach((header, index) => {
      if (header && header.toString().toLowerCase().includes('valuta')) {
        valutaColumnIndex = index;
      }
    });
    if (valutaColumnIndex >= 0 && rawJsonData.length > 1) {
      const monthYearCell = rawJsonData[1][valutaColumnIndex];
      if (monthYearCell) {
        const monthYearStr = monthYearCell.toString();
        const swedishMonths = {
          'januari': 'January',
          'februari': 'February',
          'mars': 'March',
          'april': 'April',
          'maj': 'May',
          'juni': 'June',
          'juli': 'July',
          'augusti': 'August',
          'september': 'September',
          'oktober': 'October',
          'november': 'November',
          'december': 'December'
        };
        const englishMonths = ['january','february','march','april','may','june','july','august','september','october','november','december'];
        const match = monthYearStr.match(/([a-zA-ZåäöÅÄÖ]+)\s+(\d{4})/);
        if (match) {
          const monthName = match[1].toLowerCase();
          const year = parseInt(match[2]);
          let englishMonth = swedishMonths[monthName];
          if (!englishMonth) {
            const foundMonth = englishMonths.find(m => m === monthName);
            if (foundMonth) englishMonth = foundMonth.charAt(0).toUpperCase() + foundMonth.slice(1);
          }
          if (!englishMonth) {
            for (const [swe, eng] of Object.entries(swedishMonths)) {
              if (monthName.startsWith(swe.substring(0, 3))) {
                englishMonth = eng;
                break;
              }
            }
          }
          if (englishMonth) {
            return { month: englishMonth, year: year, found: true };
          }
        }
      }
    }
    return { month: '', year: new Date().getFullYear(), found: false };
  };

  const processFile = async (uploadedFile) => {
    if (!uploadedFile) return;
    try {
      const data = await uploadedFile.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawJsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', blankrows: false, raw: false });
      if (rawJsonData.length > 0) {
        setFile(uploadedFile);
        setRawData(rawJsonData);
        setShowPreview(true);
        let maxNonEmpty = 0;
        let bestHeaderRow = 0;
        const startRow = 2;
        for (let i = startRow; i < Math.min(rawJsonData.length, 20); i++) {
          const nonEmpty = rawJsonData[i].filter(cell => cell !== '').length;
          const rowStr = rawJsonData[i].join(' ').toLowerCase();
          const hasHeaderKeywords = rowStr.includes('datum') || rowStr.includes('belopp') || rowStr.includes('mottagare') || rowStr.includes('beskrivning') || rowStr.includes('valuta');
          if (nonEmpty > maxNonEmpty || (nonEmpty === maxNonEmpty && hasHeaderKeywords)) {
            maxNonEmpty = nonEmpty;
            bestHeaderRow = i;
          }
        }
        setHeaderRow(bestHeaderRow);
        setDataStartRow(bestHeaderRow + 1);
        const { month, year, found } = extractMonthYear(rawJsonData, bestHeaderRow);
        if (found) {
          setFileMonth(month);
          setFileYear(year);
          setAutoDetectedMonth(true);
        } else {
          setFileMonth('');
          setFileYear(new Date().getFullYear());
          setAutoDetectedMonth(false);
        }
      } else {
        alert('No data found in the Excel file. Please check your file.');
      }
    } catch (error) {
      alert('Error reading file: ' + error.message + '\nPlease make sure it\'s a valid Excel file.');
      console.error('File reading error:', error);
    }
  };

  const confirmDataSelection = () => {
    if (!rawData || rawData.length === 0) return;
    const headers = rawData[headerRow].filter(h => h !== '');
    const amountCol = headers.find(h => {
      const lower = String(h).toLowerCase();
      return lower.includes('belopp') || lower.includes('amount') || lower.includes('total') || lower.includes('sum');
    });
    if (amountCol) {
      setAmountColumn(amountCol);
    }
    const processedData = [];
    for (let i = dataStartRow; i < rawData.length; i++) {
      const row = rawData[i];
      if (row.every(cell => cell === '')) continue;
      const transaction = {};
      headers.forEach((header, index) => {
        let value = row[index] || '';
        if ((String(header).toLowerCase().includes('date') || String(header).toLowerCase().includes('datum')) && value) {
          value = excelDateToJS(value);
        }
        transaction[header] = value;
      });
      if (amountCol) {
        const amountValue = transaction[amountCol];
        if (!amountValue || amountValue === '' || amountValue === 0 || String(amountValue).toLowerCase().includes('valutakurs')) {
          continue;
        }
      }
      transaction._originalIndex = processedData.length;
      processedData.push(transaction);
    }
    const dateCol = headers.find(h => String(h).toLowerCase().includes('date') || String(h).toLowerCase().includes('datum'));
    if (dateCol) {
      setDateColumn(dateCol);
    }
    let existingClassifications = {};
    let existingSplits = {};
    let wasReclassifying = false;
    headers.forEach(header => {
      if (header === 'Classification') {
        wasReclassifying = true;
        processedData.forEach((transaction, idx) => {
          if (transaction.Classification && transaction.Classification !== 'Skipped' && transaction.Classification !== 'Unclassified') {
            existingClassifications[idx] = transaction.Classification;
          }
        });
      }
      if (header === '_isSplit' || header === '_splitParent' || header === '_splitInfo') {
        processedData.forEach((transaction, idx) => {
          if (transaction._splitInfo) {
            const match = transaction._splitInfo.match(/Split \d of 2 \((\d+)%\)/);
            if (match) {
              const ratio = parseInt(match[1]);
              if (!existingSplits[idx]) {
                existingSplits[idx] = { ratio, classes: ['', ''] };
              }
            }
          }
        });
      }
    });
    setIsReclassifying(wasReclassifying);
    if (wasReclassifying) {
      setClassifications(existingClassifications);
      setSplitTransactions(existingSplits);
    }
    if (processedData.length > 0) {
      setTransactions(processedData);
      setColumns(headers.filter(h => h !== 'Classification' && h !== 'Month' && h !== 'Year' && !h.startsWith('_')));
      setCurrentIndex(0);
      if (!wasReclassifying) {
        setClassifications({});
        setSplitTransactions({});
      }
      setSkipped(new Set());
      setHistory([]);
      setShowPreview(false);
    } else {
      alert('No valid data rows found with amounts. Please check your file.');
    }
  };

  const handleFileUpload = async (event) => {
    const uploadedFile = event.target.files[0];
    await processFile(uploadedFile);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };
  const handleDrop = async (e) => {
    e.preventDefault();
    setIsDragging(false);
    const uploadedFile = e.dataTransfer.files[0];
    await processFile(uploadedFile);
  };

  const skipTransaction = () => {
    setSkipped(prev => new Set([...prev, currentIndex]));
    moveToNext();
  };

  const openSplitModal = () => {
    setShowSplitModal(true);
    setShowClassSelection(true);
    setSplitClasses({ first: '', second: '' });
    setSplitRatio(50);
  };

  const selectSplitClasses = (first, second) => {
    setSplitClasses({ first, second });
    setShowClassSelection(false);
  };

  const splitTransaction = () => {
    const transaction = transactions[currentIndex];
    if (!amountColumn || !transaction[amountColumn]) {
      alert('Cannot find amount field to split');
      return;
    }
    const originalAmount = parseAmount(transaction[amountColumn]);
    if (isNaN(originalAmount)) {
      alert('Cannot parse amount value');
      return;
    }
    const amount1 = (originalAmount * splitRatio / 100).toFixed(2);
    const amount2 = (originalAmount * (100 - splitRatio) / 100).toFixed(2);
    setSplitTransactions(prev => ({
      ...prev,
      [currentIndex]: {
        ratio: splitRatio,
        amount1,
        amount2,
        originalAmount,
        classes: [splitClasses.first, splitClasses.second]
      }
    }));
    setClassifications(prev => ({ ...prev, [currentIndex]: splitClasses.first }));
    setShowSplitModal(false);
    setSplitRatio(50);
    setSplitClasses({ first: '', second: '' });
    moveToNext();
  };

  const classifyTransaction = (classification) => {
    setHistory(prev => [...prev, { index: currentIndex, classification: classifications[currentIndex], wasSkipped: skipped.has(currentIndex), wasSplit: splitTransactions[currentIndex] }]);
    setClassifications(prev => ({ ...prev, [currentIndex]: classification }));
    if (skipped.has(currentIndex)) {
      setSkipped(prev => {
        const newSkipped = new Set(prev);
        newSkipped.delete(currentIndex);
        return newSkipped;
      });
    }
    if (splitTransactions[currentIndex]) {
      setSplitTransactions(prev => {
        const newSplits = { ...prev };
        delete newSplits[currentIndex];
        return newSplits;
      });
    }
    moveToNext();
  };

  const moveToNext = () => {
    let nextIndex = currentIndex + 1;
    while (nextIndex < transactions.length && (classifications[nextIndex] || skipped.has(nextIndex))) {
      nextIndex++;
    }
    if (nextIndex >= transactions.length) {
      const skippedIndices = Array.from(skipped).sort((a, b) => a - b);
      if (skippedIndices.length > 0) {
        nextIndex = skippedIndices[0];
      }
    }
    if (nextIndex < transactions.length) {
      setCurrentIndex(nextIndex);
    } else {
      setCurrentIndex(Math.min(currentIndex, transactions.length - 1));
    }
  };

  const undo = () => {
    if (history.length === 0) return;
    const lastAction = history[history.length - 1];
    setHistory(prev => prev.slice(0, -1));
    if (lastAction.classification === undefined) {
      setClassifications(prev => {
        const newClassifications = { ...prev };
        delete newClassifications[lastAction.index];
        return newClassifications;
      });
    } else {
      setClassifications(prev => ({ ...prev, [lastAction.index]: lastAction.classification }));
    }
    if (lastAction.wasSkipped) {
      setSkipped(prev => new Set([...prev, lastAction.index]));
    } else {
      setSkipped(prev => {
        const newSkipped = new Set(prev);
        newSkipped.delete(lastAction.index);
        return newSkipped;
      });
    }
    if (lastAction.wasSplit) {
      setSplitTransactions(prev => ({ ...prev, [lastAction.index]: lastAction.wasSplit }));
    } else {
      setSplitTransactions(prev => {
        const newSplits = { ...prev };
        delete newSplits[lastAction.index];
        return newSplits;
      });
    }
    setCurrentIndex(lastAction.index);
  };

  const navigateTransaction = (direction) => {
    if (direction === 'prev' && currentIndex > 0) {
      setCurrentIndex(currentIndex - 1);
    } else if (direction === 'next' && currentIndex < transactions.length - 1) {
      setCurrentIndex(currentIndex + 1);
    }
  };

  const calculateSubtotals = () => {
    const subtotals = { Own: 0, Family: 0, Business: 0 };
    if (!amountColumn) return subtotals;
    transactions.forEach((transaction, index) => {
      const classification = classifications[index];
      if (!classification || skipped.has(index)) return;
      let amount = parseAmount(transaction[amountColumn]);
      if (splitTransactions[index]) {
        const split = splitTransactions[index];
        if (split.classes && split.classes[0] && split.classes[1]) {
          const amt1 = parseAmount(split.amount1);
          const amt2 = parseAmount(split.amount2);
          if (subtotals.hasOwnProperty(split.classes[0])) {
            subtotals[split.classes[0]] += amt1;
          }
          if (subtotals.hasOwnProperty(split.classes[1])) {
            subtotals[split.classes[1]] += amt2;
          }
          return;
        }
      }
      if (subtotals.hasOwnProperty(classification)) {
        subtotals[classification] += amount;
      }
    });
    Object.keys(subtotals).forEach(key => {
      subtotals[key] = Math.round(subtotals[key] * 100) / 100;
    });
    return subtotals;
  };

  const exportClassifications = () => {
    let exportData = [];
    transactions.forEach((transaction, index) => {
      if (splitTransactions[index]) {
        const split = splitTransactions[index];
        const classification = classifications[index] || (skipped.has(index) ? 'Skipped' : 'Unclassified');
        const trans1 = { ...transaction, [amountColumn]: split.amount1, Classification: split.classes[0] || classification, _splitInfo: `Split 1 of 2 (${split.ratio}%)` };
        const trans2 = { ...transaction, [amountColumn]: split.amount2, Classification: split.classes[1] || classification, _splitInfo: `Split 2 of 2 (${100 - split.ratio}%)` };
        exportData.push(trans1, trans2);
      } else {
        exportData.push({ ...transaction, Classification: classifications[index] || (skipped.has(index) ? 'Skipped' : 'Unclassified') });
      }
    });
    if (dateColumn) {
      exportData.sort((a, b) => new Date(a[dateColumn]) - new Date(b[dateColumn]));
    }
    const headers = Object.keys(exportData[0] || {}).filter(k => !k.startsWith('_originalIndex'));
    const dataRows = exportData.map(row => headers.map(key => row[key]));
    const firstRow = new Array(headers.length).fill('');
    const secondRow = new Array(headers.length).fill('');
    const valutaIndex = headers.findIndex(h => h.toLowerCase().includes('valuta'));
    const monthYearIndex = valutaIndex >= 0 ? valutaIndex : 2;
    secondRow[0] = 'Månad:';
    secondRow[monthYearIndex] = `${fileMonth} ${fileYear}`;
    const sheetData = [firstRow, secondRow, headers, ...dataRows];
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Classified Purchases');
    XLSX.writeFile(workbook, `classified_purchases_${fileMonth}_${fileYear}.xlsx`);
    const classifiedOnly = exportData.filter(t => t.Classification !== 'Unclassified' && t.Classification !== 'Skipped');
    if (classifiedOnly.length > 0 && !isReclassifying) {
      const dataToSave = classifiedOnly.map(transaction => ({ ...transaction, importMonth: fileMonth, importYear: fileYear, importDate: new Date().toISOString() }));
      const stored = localStorage.getItem('purchaseClassifications');
      let existingData = [];
      if (stored) {
        try { existingData = JSON.parse(stored); } catch (e) { console.error('Error parsing stored data:', e); }
      }
      const updatedData = [...existingData, ...dataToSave];
      localStorage.setItem('purchaseClassifications', JSON.stringify(updatedData));
      setHistoricalData(updatedData);
      alert(`Exported ${exportData.length} transactions (including ${Object.keys(splitTransactions).length} split transactions) and saved ${classifiedOnly.length} classified items to database.\nTotal historical records: ${updatedData.length}`);
    } else {
      alert(`Exported ${exportData.length} transactions (including ${Object.keys(splitTransactions).length} split transactions).`);
    }
  };

  const getProgress = () => {
    const classified = Object.keys(classifications).length;
    const skippedCount = skipped.size;
    const total = transactions.length;
    const completed = classified + skippedCount;
    return { classified, skipped: skippedCount, total, completed, percentage: total > 0 ? (completed / total * 100).toFixed(1) : 0 };
  };

  const currentClassification = classifications[currentIndex];
  const isSkipped = skipped.has(currentIndex);
  const isSplit = splitTransactions[currentIndex];
  const progress = getProgress();
  const subtotals = calculateSubtotals();

  const getCompanyLogo = () => {
    if (!currentTransaction) return null;
    const companyFields = ['Mottagare', 'Vendor', 'Merchant', 'Store', 'Supplier', 'Name', 'Beskrivning'];
    for (const field of companyFields) {
      if (currentTransaction[field] && companyLogos[currentTransaction[field]]) {
        return companyLogos[currentTransaction[field]];
      }
    }
    return null;
  };

  const filteredHistorical = historicalData.filter(item => {
    if (!searchTerm) return true;
    const search = searchTerm.toLowerCase();
    return Object.values(item).some(value => String(value).toLowerCase().includes(search));
  });

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-purple-50 p-4">
      <div className="max-w-4xl mx-auto">
        {/* The component continues with the same JSX structure as provided earlier. */}
        {/* To keep the example concise, UI markup is unchanged except for parsing fixes. */}
        {/* ... UI code omitted for brevity ... */}
      </div>
    </div>
  );
};

export default PurchaseClassifier;
