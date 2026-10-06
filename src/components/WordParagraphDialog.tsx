import React from 'react';
import { X, HelpCircle } from 'lucide-react';

export interface WordParagraphDialogProps {
  show: boolean;
  onClose: () => void;
  paragraphTab: 'indents' | 'lineBreaks';
  setParagraphTab: (tab: 'indents' | 'lineBreaks') => void;
  alignment: 'left' | 'center' | 'right' | 'justify' | 'distributed';
  setAlignment: (align: 'left' | 'center' | 'right' | 'justify' | 'distributed') => void;
  outlineLevel: string;
  setOutlineLevel: (level: string) => void;
  collapsedByDefault: boolean;
  setCollapsedByDefault: (val: boolean) => void;
  indentLeft: number; // cm
  setIndentLeft: (val: number) => void;
  indentRight: number; // cm
  setIndentRight: (val: number) => void;
  specialIndent: 'none' | 'firstLine' | 'hanging';
  setSpecialIndent: (val: 'none' | 'firstLine' | 'hanging') => void;
  specialIndentBy: number; // cm
  setSpecialIndentBy: (val: number) => void;
  mirrorIndents: boolean;
  setMirrorIndents: (val: boolean) => void;
  spacingBefore: number; // pt
  setSpacingBefore: (val: number) => void;
  spacingAfter: number; // pt
  setSpacingAfter: (val: number) => void;
  lineSpacingType: 'Single' | '1.5 lines' | 'Double' | 'At least' | 'Exactly' | 'Multiple';
  setLineSpacingType: (val: 'Single' | '1.5 lines' | 'Double' | 'At least' | 'Exactly' | 'Multiple') => void;
  lineSpacingAt: number; // e.g. 0.9, 1.0, 1.15, 1.25, 1.5
  setLineSpacingAt: (val: number) => void;
  dontAddSpaceBetweenSameStyle: boolean;
  setDontAddSpaceBetweenSameStyle: (val: boolean) => void;
  onApply: () => void;
  onOpenTabsDialog?: () => void;
}

export const WordParagraphDialog: React.FC<WordParagraphDialogProps> = ({
  show,
  onClose,
  paragraphTab,
  setParagraphTab,
  alignment,
  setAlignment,
  outlineLevel,
  setOutlineLevel,
  collapsedByDefault,
  setCollapsedByDefault,
  indentLeft,
  setIndentLeft,
  indentRight,
  setIndentRight,
  specialIndent,
  setSpecialIndent,
  specialIndentBy,
  setSpecialIndentBy,
  mirrorIndents,
  setMirrorIndents,
  spacingBefore,
  setSpacingBefore,
  spacingAfter,
  setSpacingAfter,
  lineSpacingType,
  setLineSpacingType,
  lineSpacingAt,
  setLineSpacingAt,
  dontAddSpaceBetweenSameStyle,
  setDontAddSpaceBetweenSameStyle,
  onApply,
  onOpenTabsDialog,
}) => {
  if (!show) return null;

  // Calculate preview CSS styles based on settings
  const previewTextAlign: React.CSSProperties['textAlign'] =
    alignment === 'justify' ? 'justify' :
    alignment === 'center' ? 'center' :
    alignment === 'right' ? 'right' : 'left';

  const computedLineHeight =
    lineSpacingType === 'Single' ? 1.0 :
    lineSpacingType === '1.5 lines' ? 1.5 :
    lineSpacingType === 'Double' ? 2.0 :
    lineSpacingAt > 0 ? lineSpacingAt : 1.15;

  const firstLineIndentCm = specialIndent === 'firstLine' ? specialIndentBy : 0;
  const hangingIndentCm = specialIndent === 'hanging' ? -specialIndentBy : 0;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-2 sm:p-4 overflow-y-auto">
      <div
        className="w-full max-w-[480px] bg-[#f0f0f0] border border-[#a0a0a0] rounded shadow-2xl text-[12px] text-black font-sans select-none overflow-hidden my-auto"
        style={{ fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" }}
      >
        {/* Windows / Microsoft Word Classic Dialog Title Bar */}
        <div className="flex items-center justify-between px-3 py-1.5 bg-white border-b border-[#dfdfdf]">
          <span className="text-[13px] font-normal text-[#1a1a1a]">Paragraph</span>
          <div className="flex items-center gap-2 text-slate-500">
            <button
              type="button"
              className="hover:text-slate-900 p-0.5 rounded hover:bg-slate-100 transition cursor-pointer"
              title="Help"
            >
              <HelpCircle className="w-3.5 h-3.5" />
            </button>
            <button
              type="button"
              onClick={onClose}
              className="hover:text-rose-600 p-0.5 rounded hover:bg-rose-50 transition cursor-pointer"
              title="Close"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Tab Headers */}
        <div className="flex border-b border-[#c0c0c0] bg-[#f0f0f0] px-3 pt-2 gap-1">
          <button
            type="button"
            onClick={() => setParagraphTab('indents')}
            className={`px-3 py-1 text-[11.5px] border-t border-l border-r rounded-t transition cursor-pointer ${
              paragraphTab === 'indents'
                ? 'bg-white border-[#909090] text-black font-medium -mb-[1px] pb-1.5'
                : 'bg-[#e5e5e5] border-transparent text-[#404040] hover:bg-[#eaeaea]'
            }`}
          >
            <u>I</u>ndents and Spacing
          </button>
          <button
            type="button"
            onClick={() => setParagraphTab('lineBreaks')}
            className={`px-3 py-1 text-[11.5px] border-t border-l border-r rounded-t transition cursor-pointer ${
              paragraphTab === 'lineBreaks'
                ? 'bg-white border-[#909090] text-black font-medium -mb-[1px] pb-1.5'
                : 'bg-[#e5e5e5] border-transparent text-[#404040] hover:bg-[#eaeaea]'
            }`}
          >
            Line and <u>P</u>age Breaks
          </button>
        </div>

        {/* Tab Body */}
        <div className="p-4 bg-white space-y-4 border-b border-[#d0d0d0] max-h-[78vh] overflow-y-auto">
          {paragraphTab === 'indents' ? (
            <>
              {/* General Group */}
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">General</span>
                  <div className="flex-1 h-[1px] bg-[#d8d8d8]" />
                </div>

                <div className="grid grid-cols-[100px_1fr] gap-y-2 items-center pl-2">
                  <label className="text-[11.5px] text-[#222]">
                    <u>A</u>lignment:
                  </label>
                  <select
                    value={alignment}
                    onChange={(e) => setAlignment(e.target.value as any)}
                    className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 cursor-pointer shadow-inner"
                  >
                    <option value="left">Left</option>
                    <option value="center">Centered</option>
                    <option value="right">Right</option>
                    <option value="justify">Justified</option>
                    <option value="distributed">Distributed</option>
                  </select>

                  <label className="text-[11.5px] text-[#222]">
                    <u>O</u>utline level:
                  </label>
                  <div className="flex items-center gap-3">
                    <select
                      value={outlineLevel}
                      onChange={(e) => setOutlineLevel(e.target.value)}
                      className="flex-1 bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 cursor-pointer shadow-inner"
                    >
                      <option value="Body Text">Body Text</option>
                      <option value="Level 1">Level 1</option>
                      <option value="Level 2">Level 2</option>
                      <option value="Level 3">Level 3</option>
                    </select>

                    <label className="flex items-center gap-1.5 text-[11px] text-[#666] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={collapsedByDefault}
                        onChange={(e) => setCollapsedByDefault(e.target.checked)}
                        className="rounded-[2px] accent-blue-600 cursor-pointer"
                      />
                      <span>Collapsed by default</span>
                    </label>
                  </div>
                </div>
              </div>

              {/* Indentation Group */}
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">Indentation</span>
                  <div className="flex-1 h-[1px] bg-[#d8d8d8]" />
                </div>

                <div className="grid grid-cols-2 gap-4 pl-2">
                  {/* Left Column: Left & Right Indents */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        <u>L</u>eft:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="0.1"
                          min="0"
                          max="15"
                          value={indentLeft}
                          onChange={(e) => setIndentLeft(parseFloat(e.target.value) || 0)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono"
                        />
                        <span className="absolute right-6 top-1 text-[10.5px] text-gray-500 pointer-events-none">cm</span>
                      </div>
                    </div>

                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        <u>R</u>ight:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="0.1"
                          min="0"
                          max="15"
                          value={indentRight}
                          onChange={(e) => setIndentRight(parseFloat(e.target.value) || 0)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono"
                        />
                        <span className="absolute right-6 top-1 text-[10.5px] text-gray-500 pointer-events-none">cm</span>
                      </div>
                    </div>
                  </div>

                  {/* Right Column: Special & By */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        Sp<u>e</u>cial:
                      </label>
                      <select
                        value={specialIndent}
                        onChange={(e) => setSpecialIndent(e.target.value as any)}
                        className="flex-1 bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 cursor-pointer"
                      >
                        <option value="none">(none)</option>
                        <option value="firstLine">First line</option>
                        <option value="hanging">Hanging</option>
                      </select>
                    </div>

                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        <u>B</u>y:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="0.05"
                          min="0"
                          max="10"
                          disabled={specialIndent === 'none'}
                          value={specialIndentBy}
                          onChange={(e) => setSpecialIndentBy(parseFloat(e.target.value) || 0)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono disabled:bg-gray-100 disabled:text-gray-400"
                        />
                        <span className="absolute right-6 top-1 text-[10.5px] text-gray-500 pointer-events-none">cm</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="pl-2 pt-0.5">
                  <label className="flex items-center gap-1.5 text-[11.5px] text-[#222] cursor-pointer">
                    <input
                      type="checkbox"
                      checked={mirrorIndents}
                      onChange={(e) => setMirrorIndents(e.target.checked)}
                      className="rounded-[2px] accent-blue-600 cursor-pointer"
                    />
                    <span><u>M</u>irror indents</span>
                  </label>
                </div>
              </div>

              {/* Spacing Group */}
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">Spacing</span>
                  <div className="flex-1 h-[1px] bg-[#d8d8d8]" />
                </div>

                <div className="grid grid-cols-2 gap-4 pl-2">
                  {/* Left Column: Before & After (pt) */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        <u>B</u>efore:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          max="144"
                          value={spacingBefore}
                          onChange={(e) => setSpacingBefore(parseFloat(e.target.value) || 0)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono"
                        />
                        <span className="absolute right-6 top-1 text-[10.5px] text-gray-500 pointer-events-none">pt</span>
                      </div>
                    </div>

                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-12">
                        <u>A</u>fter:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          max="144"
                          value={spacingAfter}
                          onChange={(e) => setSpacingAfter(parseFloat(e.target.value) || 0)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono"
                        />
                        <span className="absolute right-6 top-1 text-[10.5px] text-gray-500 pointer-events-none">pt</span>
                      </div>
                    </div>
                  </div>

                  {/* Right Column: Line Spacing & At */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-16">
                        <u>L</u>ine spacing:
                      </label>
                      <select
                        value={lineSpacingType}
                        onChange={(e) => {
                          const val = e.target.value as any;
                          setLineSpacingType(val);
                          if (val === 'Single') setLineSpacingAt(1.0);
                          else if (val === '1.5 lines') setLineSpacingAt(1.5);
                          else if (val === 'Double') setLineSpacingAt(2.0);
                          else if (val === 'Multiple' && lineSpacingAt === 1.0) setLineSpacingAt(0.9);
                        }}
                        className="flex-1 bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 cursor-pointer"
                      >
                        <option value="Single">Single</option>
                        <option value="1.5 lines">1.5 lines</option>
                        <option value="Double">Double</option>
                        <option value="At least">At least</option>
                        <option value="Exactly">Exactly</option>
                        <option value="Multiple">Multiple</option>
                      </select>
                    </div>

                    <div className="flex items-center justify-between gap-2">
                      <label className="text-[11.5px] text-[#222] w-16">
                        <u>A</u>t:
                      </label>
                      <div className="relative flex-1">
                        <input
                          type="number"
                          step="0.05"
                          min="0.5"
                          max="5"
                          disabled={lineSpacingType === 'Single' || lineSpacingType === '1.5 lines' || lineSpacingType === 'Double'}
                          value={lineSpacingAt}
                          onChange={(e) => setLineSpacingAt(parseFloat(e.target.value) || 1)}
                          className="w-full bg-white border border-[#7a7a7a] rounded-[2px] px-2 py-1 text-[11.5px] text-black focus:outline-blue-500 font-mono disabled:bg-gray-100 disabled:text-gray-400"
                        />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="pl-2 pt-0.5">
                  <label className="flex items-center gap-1.5 text-[11.5px] text-[#222] cursor-pointer">
                    <input
                      type="checkbox"
                      checked={dontAddSpaceBetweenSameStyle}
                      onChange={(e) => setDontAddSpaceBetweenSameStyle(e.target.checked)}
                      className="rounded-[2px] accent-blue-600 cursor-pointer"
                    />
                    <span><u>D</u>on't add space between paragraphs of the same style</span>
                  </label>
                </div>
              </div>

              {/* Preview Group (Exact replica of Microsoft Word preview box) */}
              <div className="space-y-1 pt-1">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">Preview</span>
                </div>

                <div className="border border-[#7a7a7a] bg-white p-3 rounded-[1px] min-h-[110px] shadow-inner overflow-hidden select-none">
                  {/* Previous dummy paragraph */}
                  <p className="text-[8px] text-[#999999] leading-tight select-none mb-1 font-mono">
                    Previous Paragraph Previous Paragraph Previous Paragraph Previous Paragraph Previous Paragraph Previous Paragraph Previous Paragraph Previous Paragraph
                  </p>

                  {/* Active Styled Paragraph Preview with Khmer text */}
                  <div
                    style={{
                      textAlign: previewTextAlign,
                      paddingLeft: `${indentLeft * 12}px`,
                      paddingRight: `${indentRight * 12}px`,
                      textIndent: `${firstLineIndentCm * 15 + hangingIndentCm * 15}px`,
                      paddingTop: `${spacingBefore * 0.4}px`,
                      paddingBottom: `${spacingAfter * 0.4}px`,
                      lineHeight: computedLineHeight,
                    }}
                    className="text-[10px] text-black font-['Khmer_OS_Siemreap',sans-serif] transition-all duration-150"
                  >
                    ថ្ងៃទី ៣០ ខែមករា ឆ្នាំ២០២៦ ស្តុកដើមគ្រា ចំនួន ៤០.០០កេស និងចូលចំនួន ៨០.០០កេស របាយការណ៍សរុបការងារស្តុកទំនិញជាក់ស្តែង។
                  </div>

                  {/* Following dummy paragraph */}
                  <p className="text-[8px] text-[#999999] leading-tight select-none mt-1 font-mono">
                    Following Paragraph Following Paragraph Following Paragraph Following Paragraph Following Paragraph Following Paragraph Following Paragraph Following Paragraph
                  </p>
                </div>
              </div>
            </>
          ) : (
            /* Line and Page Breaks Tab */
            <div className="space-y-4 py-2">
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">Pagination</span>
                  <div className="flex-1 h-[1px] bg-[#d8d8d8]" />
                </div>
                <div className="space-y-1.5 pl-2">
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" defaultChecked className="accent-blue-600 rounded-[2px]" />
                    <span><u>W</u>idow/Orphan control</span>
                  </label>
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" className="accent-blue-600 rounded-[2px]" />
                    <span><u>K</u>eep with next</span>
                  </label>
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" className="accent-blue-600 rounded-[2px]" />
                    <span>Keep <u>l</u>ines together</span>
                  </label>
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" className="accent-blue-600 rounded-[2px]" />
                    <span><u>P</u>age break before</span>
                  </label>
                </div>
              </div>

              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className="text-[12px] font-semibold text-[#1a1a1a]">Formatting exceptions</span>
                  <div className="flex-1 h-[1px] bg-[#d8d8d8]" />
                </div>
                <div className="space-y-1.5 pl-2">
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" className="accent-blue-600 rounded-[2px]" />
                    <span><u>S</u>uppress line numbers</span>
                  </label>
                  <label className="flex items-center gap-2 text-[11.5px] text-[#222] cursor-pointer">
                    <input type="checkbox" className="accent-blue-600 rounded-[2px]" />
                    <span><u>D</u>on't hyphenate</span>
                  </label>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer Action Buttons */}
        <div className="flex items-center justify-between p-3 bg-[#f0f0f0]">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onOpenTabsDialog}
              className="px-3 py-1 bg-[#e1e1e1] hover:bg-[#d5d5d5] active:bg-[#c8c8c8] text-black border border-[#7a7a7a] rounded-[2px] text-[11.5px] transition cursor-pointer shadow-xs"
            >
              <u>T</u>abs...
            </button>
            <button
              type="button"
              onClick={() => {
                alert('Saved as default paragraph styling template.');
              }}
              className="px-3 py-1 bg-[#e1e1e1] hover:bg-[#d5d5d5] active:bg-[#c8c8c8] text-black border border-[#7a7a7a] rounded-[2px] text-[11.5px] transition cursor-pointer shadow-xs"
            >
              Set As Defa<u>u</u>lt
            </button>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                onApply();
                onClose();
              }}
              className="px-4 py-1 bg-[#e1e1e1] hover:bg-[#0078d7] hover:text-white active:bg-[#005a9e] text-black border border-[#0078d7] rounded-[2px] text-[11.5px] font-medium transition cursor-pointer shadow-xs min-w-[70px]"
            >
              OK
            </button>
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1 bg-[#e1e1e1] hover:bg-[#d5d5d5] active:bg-[#c8c8c8] text-black border border-[#7a7a7a] rounded-[2px] text-[11.5px] transition cursor-pointer shadow-xs min-w-[70px]"
            >
              Cancel
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
