import React, { useState, useRef, useEffect, useCallback } from 'react';
import { createPortal } from 'react-dom';
import { ChevronsUpDown, ChevronDown, Calendar as CalendarIcon } from 'lucide-react';

export interface CustomDatePickerProps {
  value: string; // YYYY-MM-DD
  onChange: (dateStr: string) => void;
  placeholder?: string;
  className?: string;
  disabled?: boolean;
  required?: boolean;
  id?: string;
  name?: string;
  title?: string;
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

export const CustomDatePicker: React.FC<CustomDatePickerProps> = ({
  value,
  onChange,
  placeholder = 'YYYY-MM-DD',
  className = '',
  disabled = false,
  required = false,
  id,
  name,
  title,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [coords, setCoords] = useState<{ top: number; left: number }>({
    top: 0,
    left: 0,
  });

  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);

  // Parse safe date
  const parseDateSafe = (val?: string) => {
    if (!val) return new Date();
    const parts = val.trim().split(/[-/]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[2], 10);
        return new Date(y, m, d);
      } else if (parts[2].length === 4) {
        // DD/MM/YYYY or MM/DD/YYYY
        const y = parseInt(parts[2], 10);
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[0], 10);
        return new Date(y, m, d);
      }
    }
    const parsed = new Date(val);
    return isNaN(parsed.getTime()) ? new Date() : parsed;
  };

  const parsedInitial = parseDateSafe(value);
  const [viewYear, setViewYear] = useState<number>(parsedInitial.getFullYear());
  const [viewMonth, setViewMonth] = useState<number>(parsedInitial.getMonth()); // 0-11

  // Synchronize view month/year when value changes
  useEffect(() => {
    if (value) {
      const d = parseDateSafe(value);
      if (!isNaN(d.getTime())) {
        setViewYear(d.getFullYear());
        setViewMonth(d.getMonth());
      }
    }
  }, [value]);

  // Recalculate popup position relative to viewport using Portal & Fixed Position
  const updatePosition = useCallback(() => {
    if (!inputRef.current) return;
    const rect = inputRef.current.getBoundingClientRect();
    const popoverHeight = 350; // approximate height of popover
    const popoverWidth = 292;  // width of popover

    const viewportHeight = window.innerHeight;
    const viewportWidth = window.innerWidth;

    const spaceBelow = viewportHeight - rect.bottom;
    const spaceAbove = rect.top;

    // Determine placement: top if space below is less than popover height AND space above is larger
    const openAbove = spaceBelow < popoverHeight && spaceAbove > spaceBelow;

    let top = openAbove ? rect.top - popoverHeight - 6 : rect.bottom + 6;
    let left = rect.left;

    // Prevent horizontal overflow right
    if (left + popoverWidth > viewportWidth - 12) {
      left = Math.max(12, viewportWidth - popoverWidth - 12);
    }
    // Prevent horizontal overflow left
    if (left < 12) {
      left = 12;
    }

    // Prevent top overflow
    if (top < 12) {
      top = 12;
    }

    setCoords({
      top,
      left,
    });
  }, []);

  // Update position on open, scroll, or window resize
  useEffect(() => {
    if (isOpen) {
      updatePosition();

      const handleScrollOrResize = () => {
        updatePosition();
      };

      window.addEventListener('scroll', handleScrollOrResize, true);
      window.addEventListener('resize', handleScrollOrResize);

      return () => {
        window.removeEventListener('scroll', handleScrollOrResize, true);
        window.removeEventListener('resize', handleScrollOrResize);
      };
    }
  }, [isOpen, updatePosition]);

  // Close on outside click
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      const target = e.target as Node;
      if (
        containerRef.current &&
        !containerRef.current.contains(target) &&
        popoverRef.current &&
        !popoverRef.current.contains(target)
      ) {
        setIsOpen(false);
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);

  const formatDateToISO = (y: number, m: number, d: number) => {
    return `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  };

  // Calendar calculations
  const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
  const firstDayOfWeek = new Date(viewYear, viewMonth, 1).getDay(); // 0 = Sunday
  const prevMonthDays = new Date(viewYear, viewMonth, 0).getDate();

  const daysGrid: Array<{ day: number; month: 'prev' | 'current' | 'next'; dateStr: string }> = [];

  // Trailing previous month days
  for (let i = firstDayOfWeek - 1; i >= 0; i--) {
    const d = prevMonthDays - i;
    const pMonth = viewMonth === 0 ? 11 : viewMonth - 1;
    const pYear = viewMonth === 0 ? viewYear - 1 : viewYear;
    daysGrid.push({ day: d, month: 'prev', dateStr: formatDateToISO(pYear, pMonth, d) });
  }

  // Current month days
  for (let d = 1; d <= daysInMonth; d++) {
    daysGrid.push({ day: d, month: 'current', dateStr: formatDateToISO(viewYear, viewMonth, d) });
  }

  // Leading next month days
  const totalCells = daysGrid.length <= 35 ? 35 : 42;
  const nextDaysCount = totalCells - daysGrid.length;
  for (let d = 1; d <= nextDaysCount; d++) {
    const nMonth = viewMonth === 11 ? 0 : viewMonth + 1;
    const nYear = viewMonth === 11 ? viewYear + 1 : viewYear;
    daysGrid.push({ day: d, month: 'next', dateStr: formatDateToISO(nYear, nMonth, d) });
  }

  const yearsList: number[] = [];
  for (let y = 1950; y <= 2040; y++) {
    yearsList.push(y);
  }

  const today = new Date();
  const todayStr = formatDateToISO(today.getFullYear(), today.getMonth(), today.getDate());

  const renderPopover = () => {
    if (!isOpen) return null;

    const popoverContent = (
      <div
        ref={popoverRef}
        style={{
          position: 'fixed',
          top: `${coords.top}px`,
          left: `${coords.left}px`,
          zIndex: 99999,
        }}
        className="w-[292px] bg-white border border-gray-200/90 rounded-2xl shadow-2xl p-4 font-sans select-none animate-in fade-in-50 zoom-in-95 duration-100"
      >
        {/* Header Selectors: Month (Left) & Year (Right) */}
        <div className="flex items-center gap-2 mb-1">
          {/* Month Select Box */}
          <div className="relative flex-1">
            <select
              value={viewMonth}
              onChange={(e) => setViewMonth(parseInt(e.target.value, 10))}
              className="w-full appearance-none bg-white border border-gray-200 hover:border-gray-300 rounded-xl px-3 py-2 pr-7 text-sm font-normal text-gray-800 focus:outline-none focus:ring-1 focus:ring-gray-300 cursor-pointer shadow-2xs transition"
            >
              {MONTHS.map((m, idx) => (
                <option key={m} value={idx}>
                  {m}
                </option>
              ))}
            </select>
            <ChevronsUpDown className="w-3.5 h-3.5 text-gray-400 absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
          </div>

          {/* Year Select Box */}
          <div className="relative w-28">
            <select
              value={viewYear}
              onChange={(e) => setViewYear(parseInt(e.target.value, 10))}
              className="w-full appearance-none bg-white border border-gray-200 hover:border-gray-300 rounded-xl px-3 py-2 pr-7 text-sm font-normal text-gray-800 focus:outline-none focus:ring-1 focus:ring-gray-300 cursor-pointer shadow-2xs transition"
            >
              {yearsList.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
            <ChevronDown className="w-4 h-4 text-gray-400 absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
          </div>
        </div>

        {/* Weekday Row */}
        <div className="grid grid-cols-7 gap-1 text-center font-normal text-gray-400 text-xs mt-3 mb-2">
          <span>Su</span>
          <span>Mo</span>
          <span>Tu</span>
          <span>We</span>
          <span>Th</span>
          <span>Fr</span>
          <span>Sa</span>
        </div>

        {/* Days Grid */}
        <div className="grid grid-cols-7 gap-y-1.5 gap-x-1 text-center">
          {daysGrid.map((item, idx) => {
            const isSelected = item.dateStr === value;
            const isToday = item.dateStr === todayStr;
            const isCurrentMonth = item.month === 'current';

            return (
              <button
                key={idx}
                type="button"
                onClick={() => {
                  onChange(item.dateStr);
                  setIsOpen(false);
                }}
                className={`w-9 h-9 mx-auto flex items-center justify-center rounded-xl text-sm transition cursor-pointer ${
                  isSelected
                    ? 'bg-[#18181b] text-white font-medium shadow-xs'
                    : isToday
                    ? 'bg-gray-100 text-gray-900 font-medium hover:bg-gray-200'
                    : isCurrentMonth
                    ? 'text-gray-800 hover:bg-gray-100 font-normal'
                    : 'text-gray-300 hover:bg-gray-50 font-normal'
                }`}
              >
                {item.day}
              </button>
            );
          })}
        </div>
      </div>
    );

    return createPortal(popoverContent, document.body);
  };

  return (
    <div className="relative inline-block w-full text-left" ref={containerRef}>
      {/* Input box */}
      <div className="relative flex items-center w-full">
        <input
          ref={inputRef}
          type="text"
          id={id}
          name={name}
          title={title}
          disabled={disabled}
          required={required}
          value={value || ''}
          placeholder={placeholder}
          onClick={() => !disabled && setIsOpen(true)}
          onChange={(e) => onChange(e.target.value)}
          className={`w-full bg-white border border-gray-300 rounded px-2.5 py-1 text-xs text-gray-800 placeholder-gray-400 focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 cursor-pointer pr-7 transition-colors shadow-2xs ${className}`}
        />
        <button
          type="button"
          tabIndex={-1}
          disabled={disabled}
          onClick={(e) => {
            e.stopPropagation();
            if (!disabled) {
              setIsOpen(!isOpen);
            }
          }}
          className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 p-0.5 cursor-pointer focus:outline-none flex items-center justify-center"
          title="ជ្រើសរើសកាលបរិច្ឆេទ"
        >
          <CalendarIcon className="w-3.5 h-3.5 shrink-0" />
        </button>
      </div>

      {renderPopover()}
    </div>
  );
};
