import React, { useRef, useState, useEffect } from 'react';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';

interface StandardCanvasContainerProps {
  children: React.ReactNode;
  contentWidth?: number; // default 1122 for landscape A4 or 794 for portrait
  className?: string;
  enableAutoFit?: boolean;
}

export const StandardCanvasContainer: React.FC<StandardCanvasContainerProps> = ({
  children,
  contentWidth = 1122,
  className = '',
  enableAutoFit = true,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const { scaleMode } = useWorkspaceSettings();
  const [autoScale, setAutoScale] = useState<number>(1);

  useEffect(() => {
    if (!enableAutoFit) return;

    const calculateScale = () => {
      if (!containerRef.current) return;
      const containerWidth = containerRef.current.clientWidth - 24; // padding allowance
      if (scaleMode === 'fit') {
        const ratio = Math.min(1.05, Math.max(0.45, containerWidth / contentWidth));
        setAutoScale(ratio);
      } else {
        const numeric = parseInt(scaleMode.replace('%', ''), 10) / 100;
        setAutoScale(isNaN(numeric) ? 1 : numeric);
      }
    };

    calculateScale();

    const resizeObserver = new ResizeObserver(() => {
      calculateScale();
    });

    if (containerRef.current) {
      resizeObserver.observe(containerRef.current);
    }

    window.addEventListener('resize', calculateScale);

    return () => {
      resizeObserver.disconnect();
      window.removeEventListener('resize', calculateScale);
    };
  }, [scaleMode, contentWidth, enableAutoFit]);

  return (
    <div
      ref={containerRef}
      className={`w-full overflow-x-auto flex justify-center print:overflow-visible print:block ${className}`}
    >
      <div
        style={{
          transform: autoScale !== 1 ? `scale(${autoScale})` : undefined,
          transformOrigin: 'top center',
          transition: 'transform 0.15s ease-out',
        }}
        className="shrink-0 print:transform-none"
      >
        {children}
      </div>
    </div>
  );
};
