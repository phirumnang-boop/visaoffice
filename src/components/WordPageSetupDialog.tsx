import React from 'react';
import { Layout, X } from 'lucide-react';

export interface CustomMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface WordPageSetupDialogProps {
  show: boolean;
  onClose: () => void;
  pageSetupTab: 'margins' | 'paper' | 'layout';
  setPageSetupTab: (tab: 'margins' | 'paper' | 'layout') => void;
  tempMargins: CustomMargins;
  setTempMargins: React.Dispatch<React.SetStateAction<CustomMargins>>;
  gutter: number;
  setGutter: (val: number) => void;
  gutterPosition: 'left' | 'top';
  setGutterPosition: (val: 'left' | 'top') => void;
  pageOrientation: 'portrait' | 'landscape';
  setPageOrientation: (val: 'portrait' | 'landscape') => void;
  paperSize: string;
  setPaperSize: (val: string) => void;
  headerMargin: number;
  setHeaderMargin: (val: number) => void;
  footerMargin: number;
  setFooterMargin: (val: number) => void;
  activeTabStop: number;
  setActiveTabStop: (val: number) => void;
  onApply: (margins: CustomMargins) => void;
}

export const WordPageSetupDialog: React.FC<WordPageSetupDialogProps> = ({
  show,
  onClose,
  pageSetupTab,
  setPageSetupTab,
  tempMargins,
  setTempMargins,
  gutter,
  setGutter,
  gutterPosition,
  setGutterPosition,
  pageOrientation,
  setPageOrientation,
  paperSize,
  setPaperSize,
  headerMargin,
  setHeaderMargin,
  footerMargin,
  setFooterMargin,
  activeTabStop,
  setActiveTabStop,
  onApply,
}) => {
  if (!show) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/65 backdrop-blur-[2px] flex items-center justify-center p-3 animate-fadeIn">
      <div className="bg-[#F0F0F0] text-gray-900 rounded shadow-2xl border border-gray-400 w-full max-w-[530px] text-[11px] select-none font-sans overflow-hidden">
        {/* Window Header */}
        <div className="flex items-center justify-between bg-gradient-to-r from-blue-800 to-indigo-900 text-white px-3 py-1.5 font-semibold text-xs border-b border-blue-900">
          <span className="flex items-center gap-1.5">
            <Layout className="w-3.5 h-3.5 text-blue-300" />
            <span>Page Setup (កំណត់ទំព័រ និងគែមក្រដាស A4)</span>
          </span>
          <button
            type="button"
            onClick={onClose}
            className="hover:bg-red-600 text-white p-0.5 rounded transition cursor-pointer"
            title="បិទ"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Tab Navigation */}
        <div className="flex items-center bg-[#E5E5E5] px-3 pt-2 border-b border-gray-300 gap-1 text-xs">
          <button
            type="button"
            onClick={() => setPageSetupTab('margins')}
            className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
              pageSetupTab === 'margins'
                ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                : 'bg-transparent border-transparent text-gray-600 hover:text-black'
            }`}
          >
            Margins (គែមក្រដាស)
          </button>
          <button
            type="button"
            onClick={() => setPageSetupTab('paper')}
            className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
              pageSetupTab === 'paper'
                ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                : 'bg-transparent border-transparent text-gray-600 hover:text-black'
            }`}
          >
            Paper (ទំហំក្រដាស)
          </button>
          <button
            type="button"
            onClick={() => setPageSetupTab('layout')}
            className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
              pageSetupTab === 'layout'
                ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                : 'bg-transparent border-transparent text-gray-600 hover:text-black'
            }`}
          >
            Layout (ប្លង់ & គម្លាត)
          </button>
        </div>

        {/* Tab Content Body */}
        <div className="p-3.5 space-y-3 bg-[#F0F0F0]">
          {pageSetupTab === 'margins' && (
            <div className="space-y-3">
              {/* Margins Inputs Box */}
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Margins (គែមក្រដាសគិតជា សង់ទីម៉ែត្រ cm)
                </legend>
                <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-xs">
                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Top (លើ) :</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.1"
                        min="0.2"
                        max="5.0"
                        value={tempMargins.top}
                        onChange={(e) => setTempMargins({ ...tempMargins, top: parseFloat(e.target.value) || 0 })}
                        className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Bottom (ក្រោម) :</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.1"
                        min="0.2"
                        max="5.0"
                        value={tempMargins.bottom}
                        onChange={(e) => setTempMargins({ ...tempMargins, bottom: parseFloat(e.target.value) || 0 })}
                        className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-gray-900 font-bold">Left (ឆ្វេង) :</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.1"
                        min="0.2"
                        max="6.0"
                        value={tempMargins.left}
                        onChange={(e) => setTempMargins({ ...tempMargins, left: parseFloat(e.target.value) || 0 })}
                        className="w-16 bg-blue-50 border border-blue-400 font-bold rounded px-1.5 py-0.5 text-right font-mono text-blue-900"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-gray-900 font-bold">Right (ស្តាំ) :</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.1"
                        min="0.2"
                        max="5.0"
                        value={tempMargins.right}
                        onChange={(e) => setTempMargins({ ...tempMargins, right: parseFloat(e.target.value) || 0 })}
                        className="w-16 bg-blue-50 border border-blue-400 font-bold rounded px-1.5 py-0.5 text-right font-mono text-blue-900"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Gutter :</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.1"
                        min="0"
                        max="3.0"
                        value={gutter}
                        onChange={(e) => setGutter(parseFloat(e.target.value) || 0)}
                        className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Gutter pos :</label>
                    <select
                      value={gutterPosition}
                      onChange={(e) => setGutterPosition(e.target.value as any)}
                      className="w-20 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black"
                    >
                      <option value="left">Left</option>
                      <option value="top">Top</option>
                    </select>
                  </div>
                </div>
              </fieldset>

              {/* Orientation */}
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Orientation (ទិសដៅក្រដាស)
                </legend>
                <div className="flex items-center gap-8 justify-center py-1">
                  <label
                    onClick={() => setPageOrientation('portrait')}
                    className={`flex flex-col items-center gap-1.5 p-2 rounded cursor-pointer border transition ${
                      pageOrientation === 'portrait'
                        ? 'border-blue-600 bg-blue-50 font-bold text-blue-900'
                        : 'border-transparent hover:bg-gray-100 text-gray-700'
                    }`}
                  >
                    <div className="w-7 h-10 border-2 border-gray-700 bg-white shadow-xs rounded-xs flex flex-col justify-between p-0.5">
                      <div className="w-full h-1 bg-gray-400"></div>
                      <div className="w-full h-0.5 bg-gray-300"></div>
                      <div className="w-full h-0.5 bg-gray-300"></div>
                    </div>
                    <span>Portrait (បញ្ឈរ)</span>
                  </label>

                  <label
                    onClick={() => setPageOrientation('landscape')}
                    className={`flex flex-col items-center gap-1.5 p-2 rounded cursor-pointer border transition ${
                      pageOrientation === 'landscape'
                        ? 'border-blue-600 bg-blue-50 font-bold text-blue-900'
                        : 'border-transparent hover:bg-gray-100 text-gray-700'
                    }`}
                  >
                    <div className="w-10 h-7 border-2 border-gray-700 bg-white shadow-xs rounded-xs flex flex-col justify-between p-0.5">
                      <div className="w-full h-1 bg-gray-400"></div>
                      <div className="w-full h-0.5 bg-gray-300"></div>
                      <div className="w-full h-0.5 bg-gray-300"></div>
                    </div>
                    <span>Landscape (ផ្តេក)</span>
                  </label>
                </div>
              </fieldset>
            </div>
          )}

          {pageSetupTab === 'paper' && (
            <div className="space-y-3">
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Paper size (ទំហំក្រដាសបោះពុម្ព)
                </legend>
                <div className="space-y-2 text-xs">
                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Paper size:</label>
                    <select
                      value={paperSize}
                      onChange={(e) => setPaperSize(e.target.value)}
                      className="bg-white border border-gray-300 rounded px-2 py-1 text-black font-semibold w-48"
                    >
                      <option value="A4">A4 (21.0 cm x 29.7 cm)</option>
                      <option value="Letter">Letter (21.59 cm x 27.94 cm)</option>
                      <option value="Legal">Legal (21.59 cm x 35.56 cm)</option>
                      <option value="A3">A3 (29.7 cm x 42.0 cm)</option>
                    </select>
                  </div>
                  <div className="grid grid-cols-2 gap-2 pt-2 text-gray-600 border-t border-gray-200">
                    <div>Width (ទទឹង) : <span className="font-mono font-bold text-gray-900">21.0 cm</span></div>
                    <div>Height (កម្ពស់) : <span className="font-mono font-bold text-gray-900">29.7 cm</span></div>
                  </div>
                </div>
              </fieldset>
            </div>
          )}

          {pageSetupTab === 'layout' && (
            <div className="space-y-3">
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Headers and Footers (ក្បាល និងជើងទំព័រ)
                </legend>
                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Header (ក្បាល):</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.05"
                        min="0"
                        max="5"
                        value={headerMargin}
                        onChange={(e) => setHeaderMargin(parseFloat(e.target.value) || 0)}
                        className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>
                  <div className="flex items-center justify-between">
                    <label className="text-gray-700">Footer (ជើង):</label>
                    <div className="flex items-center">
                      <input
                        type="number"
                        step="0.05"
                        min="0"
                        max="5"
                        value={footerMargin}
                        onChange={(e) => setFooterMargin(parseFloat(e.target.value) || 0)}
                        className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                      />
                      <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                    </div>
                  </div>
                </div>
              </fieldset>

              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Tabs & Paragraph Alignment (គម្លាតកថាខណ្ឌ & Tab Stop)
                </legend>
                <div className="flex items-center justify-between text-xs">
                  <label className="text-gray-700">Default Tab Stop (គម្លាតដើមបន្ទាត់ Tab) :</label>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="number"
                      step="0.1"
                      min="0"
                      max="4"
                      value={activeTabStop}
                      onChange={(e) => setActiveTabStop(parseFloat(e.target.value) || 0)}
                      className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                    />
                    <span className="text-gray-500 text-[10px]">cm</span>
                    <button
                      type="button"
                      onClick={() => setActiveTabStop(1.2)}
                      className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded text-[10px] cursor-pointer"
                    >
                      1.2cm
                    </button>
                  </div>
                </div>
              </fieldset>
            </div>
          )}

          {/* Live Preview Box */}
          <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
            <legend className="text-gray-700 px-1 font-semibold text-[11px]">
              Preview (ទិដ្ឋភាពជាក់ស្តែងនៃទំព័រ)
            </legend>
            <div className="flex items-center justify-center p-2 bg-white border border-gray-300 rounded shadow-inner">
              {/* Scaled Visual Page representation */}
              <div
                style={{
                  width: pageOrientation === 'portrait' ? '120px' : '170px',
                  height: pageOrientation === 'portrait' ? '170px' : '120px',
                  paddingTop: `${Math.max(2, (tempMargins.top / 29.7) * 170)}px`,
                  paddingBottom: `${Math.max(2, (tempMargins.bottom / 29.7) * 170)}px`,
                  paddingLeft: `${Math.max(2, (tempMargins.left / 21.0) * 120)}px`,
                  paddingRight: `${Math.max(2, (tempMargins.right / 21.0) * 120)}px`,
                }}
                className="border-2 border-gray-700 bg-white shadow flex flex-col justify-between transition-all relative overflow-hidden"
              >
                {/* Top simulated lines */}
                <div className="space-y-0.5 border border-dashed border-blue-300 p-0.5 bg-blue-50/50 rounded-xs h-full flex flex-col justify-between">
                  <div className="flex justify-between items-center">
                    <div className="w-1/3 h-1 bg-gray-400 rounded-xs"></div>
                    <div className="w-1/3 h-1 bg-gray-400 rounded-xs"></div>
                  </div>
                  <div className="space-y-0.5 py-1">
                    <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                    <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                    <div className="w-3/4 h-0.5 bg-gray-300 rounded-xs"></div>
                    <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                  </div>
                  <div className="flex justify-between items-center pt-1">
                    <div className="w-1/4 h-1 bg-gray-400 rounded-xs"></div>
                    <div className="w-1/4 h-1 bg-gray-400 rounded-xs"></div>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex items-center justify-between pt-2 text-[10.5px]">
              <span className="text-gray-600">Apply to :</span>
              <select className="bg-white border border-gray-300 rounded px-2 py-0.5 text-black">
                <option>Whole document (ឯកសារទាំងមូល)</option>
                <option>This point forward</option>
              </select>
            </div>
          </fieldset>
        </div>

        {/* Bottom Actions */}
        <div className="flex items-center justify-between px-3.5 py-2.5 bg-[#E8E8E8] border-t border-gray-300">
          <button
            type="button"
            onClick={() => {
              onApply(tempMargins);
              alert('បានកំណត់គែមក្រដាសជាលំនាំដើម (Default Margins Saved)!');
            }}
            className="px-2.5 py-1 bg-white hover:bg-gray-50 text-gray-800 rounded border border-gray-300 font-medium text-[11px] shadow-sm cursor-pointer transition"
          >
            Set As Default
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                onApply(tempMargins);
                onClose();
              }}
              className="px-5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded border border-blue-700 font-semibold text-xs shadow-sm cursor-pointer transition"
            >
              OK (យល់ព្រម)
            </button>
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1 bg-white hover:bg-gray-100 text-gray-800 rounded border border-gray-400 font-medium text-xs shadow-sm cursor-pointer transition"
            >
              Cancel (បោះបង់)
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
