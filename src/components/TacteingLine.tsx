import React, { useState, useEffect } from 'react';
import { Upload, Image as ImageIcon, RotateCcw, Check, Sparkles } from 'lucide-react';

export type TacteingType = 'image2-classic' | 'symbol-cross' | 'symbol-flower' | 'custom-image';

interface TacteingLineProps {
  type?: TacteingType;
  customImage?: string | null;
  width?: number; // e.g. 140
  height?: number; // e.g. 14
  className?: string;
}

const STORAGE_KEY_TYPE = 'app_tacteing_type';
const STORAGE_KEY_IMG = 'app_tacteing_custom_img';

export const getSavedTacteingSettings = (): { type: TacteingType; customImage: string | null } => {
  try {
    const savedType = localStorage.getItem(STORAGE_KEY_TYPE) as TacteingType | null;
    const savedImg = localStorage.getItem(STORAGE_KEY_IMG);
    return {
      type: savedType || 'image2-classic',
      customImage: savedImg || null,
    };
  } catch (e) {
    return { type: 'image2-classic', customImage: null };
  }
};

export const saveTacteingSettings = (type: TacteingType, customImage?: string | null) => {
  try {
    localStorage.setItem(STORAGE_KEY_TYPE, type);
    if (customImage) {
      localStorage.setItem(STORAGE_KEY_IMG, customImage);
    } else if (type !== 'custom-image') {
      localStorage.removeItem(STORAGE_KEY_IMG);
    }
    window.dispatchEvent(new Event('tacteing_settings_updated'));
  } catch (e) {
    console.error('Failed to save tacteing settings:', e);
  }
};

export const TacteingLine: React.FC<TacteingLineProps> = ({
  type: propType,
  customImage: propCustomImage,
  width = 160,
  height = 16,
  className = '',
}) => {
  const [settings, setSettings] = useState<{ type: TacteingType; customImage: string | null }>(
    getSavedTacteingSettings()
  );

  useEffect(() => {
    const updateSettings = () => {
      setSettings(getSavedTacteingSettings());
    };
    window.addEventListener('tacteing_settings_updated', updateSettings);
    window.addEventListener('storage', updateSettings);
    return () => {
      window.removeEventListener('tacteing_settings_updated', updateSettings);
      window.removeEventListener('storage', updateSettings);
    };
  }, []);

  const activeType = propType !== undefined ? propType : settings.type;
  const activeImg = propCustomImage !== undefined ? propCustomImage : settings.customImage;

  // If custom image is provided or set in localStorage
  if (activeType === 'custom-image' && activeImg) {
    return (
      <div className={`flex items-center justify-center pt-0 mt-0.5 mb-0.5 ${className}`} style={{ lineHeight: 1 }}>
        <img
          src={activeImg}
          alt="តាក់តែង"
          style={{ height: `${height}px`, maxWidth: `${width + 40}px`, objectFit: 'contain', imageRendering: 'high-quality' }}
          className="mx-auto select-none print:max-h-6"
        />
      </div>
    );
  }

  if (activeType === 'symbol-cross') {
    return (
      <div className={`flex items-center justify-center pt-0 mt-0.5 mb-0.5 text-black ${className}`} style={{ lineHeight: 1, height: `${height}px` }}>
        <div className="flex items-center gap-1 font-mono text-[13px] tracking-tighter select-none font-bold">
          <span className="h-[2px] w-14 bg-black inline-block"></span>
          <span>(</span>
          <span className="text-[14px] font-bold">+</span>
          <span>)</span>
          <span className="h-[2px] w-14 bg-black inline-block"></span>
        </div>
      </div>
    );
  }

  if (activeType === 'symbol-flower') {
    return (
      <div className={`flex items-center justify-center pt-0 mt-0.5 mb-0.5 text-black ${className}`} style={{ lineHeight: 1, height: `${height}px` }}>
        <div className="flex items-center gap-1 font-mono text-[13px] tracking-tighter select-none font-bold">
          <span className="h-[2px] w-14 bg-black inline-block"></span>
          <span>(</span>
          <span className="text-[13px]">❁</span>
          <span>)</span>
          <span className="h-[2px] w-14 bg-black inline-block"></span>
        </div>
      </div>
    );
  }

  // Default: Classical Image 2 ornament SVG vector reproduction (Crisp, High Quality, Deep Pure Black)
  return (
    <div className={`flex items-center justify-center pt-0 mt-0.5 mb-0.5 ${className}`} style={{ lineHeight: 1 }}>
      <svg
        style={{ height: `${height || 16}px`, width: `${width || 160}px` }}
        className="text-black fill-black select-none"
        viewBox="0 0 240 24"
        preserveAspectRatio="xMidYMid meet"
        shapeRendering="geometricPrecision"
      >
        <g id="classical-khmer-tacteing" stroke="#000000" fill="#000000">
          {/* Tapered Left Line */}
          <path d="M 2 12 L 76 10.5 L 76 13.5 Z" stroke="#000000" strokeWidth="0.8" fill="#000000" />

          {/* Left Wing Scroll & Loop */}
          <path d="M 76 12 C 80 4.5, 92 3.5, 96 12 C 92 20.5, 80 19.5, 76 12 Z" fill="none" stroke="#000000" strokeWidth="2.8" />
          <circle cx="86" cy="12" r="3" fill="#000000" />
          <path d="M 91 12 C 89 5.5, 95 2, 93 1 C 90 1, 88 5.5, 90 12 C 88 18.5, 90 23, 93 23 C 95 22, 89 18.5, 91 12 Z" fill="#000000" />

          {/* Center Flower Ornament */}
          <circle cx="120" cy="12" r="3.2" fill="#000000" />
          {/* 8 Flower Petals */}
          <path d="M 120 7.5 C 117 3.5, 117 0.5, 120 -1 C 123 0.5, 123 3.5, 120 7.5 Z" fill="#000000" />
          <path d="M 120 16.5 C 117 20.5, 117 23.5, 120 25 C 123 23.5, 123 20.5, 120 16.5 Z" fill="#000000" />
          <path d="M 115.5 12 C 111.5 9, 108.5 9, 107 12 C 108.5 15, 111.5 15, 115.5 12 Z" fill="#000000" />
          <path d="M 124.5 12 C 128.5 9, 131.5 9, 133 12 C 131.5 15, 128.5 15, 124.5 12 Z" fill="#000000" />
          
          <path d="M 117.8 9.8 C 114.5 6.5, 112 4.5, 111 5.8 C 112.5 7.5, 114.5 9.2, 117.8 9.8 Z" fill="#000000" />
          <path d="M 122.2 9.8 C 125.5 6.5, 128 4.5, 129 5.8 C 127.5 7.5, 125.5 9.2, 122.2 9.8 Z" fill="#000000" />
          <path d="M 117.8 14.2 C 114.5 17.5, 112 19.5, 111 18.2 C 112.5 16.5, 114.5 14.8, 117.8 14.2 Z" fill="#000000" />
          <path d="M 122.2 14.2 C 125.5 17.5, 128 19.5, 129 18.2 C 127.5 16.5, 125.5 14.8, 122.2 14.2 Z" fill="#000000" />

          {/* Right Wing Scroll & Loop */}
          <path d="M 164 12 C 160 4.5, 148 3.5, 144 12 C 148 20.5, 160 19.5, 164 12 Z" fill="none" stroke="#000000" strokeWidth="2.8" />
          <circle cx="154" cy="12" r="3" fill="#000000" />
          <path d="M 149 12 C 151 5.5, 145 2, 147 1 C 150 1, 152 5.5, 150 12 C 152 18.5, 150 23, 147 23 C 145 22, 151 18.5, 149 12 Z" fill="#000000" />

          {/* Tapered Right Line */}
          <path d="M 238 12 L 164 10.5 L 164 13.5 Z" stroke="#000000" strokeWidth="0.8" fill="#000000" />
        </g>
      </svg>
    </div>
  );
};

interface TacteingControlsProps {
  currentType: TacteingType;
  customImage: string | null;
  onChange: (type: TacteingType, customImg?: string | null) => void;
}

export const TacteingControlSelector: React.FC<TacteingControlsProps> = ({
  currentType,
  customImage,
  onChange,
}) => {
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        if (evt.target?.result) {
          const dataUrl = evt.target.result as string;
          onChange('custom-image', dataUrl);
          saveTacteingSettings('custom-image', dataUrl);
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleSelectType = (type: TacteingType) => {
    onChange(type, customImage);
    saveTacteingSettings(type, customImage);
  };

  return (
    <div className="bg-amber-50/70 border border-amber-200 rounded-md p-3 text-xs space-y-2.5 my-2">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 font-bold text-amber-900">
          <Sparkles className="w-3.5 h-3.5 text-amber-600" />
          <span>ជ្រើសរើស ឬបញ្ចូលរូបភាព តាក់តែង (TacTeing Style)</span>
        </div>
        <span className="text-[10px] text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded font-mono">
          លិខិតផ្លូវការ
        </span>
      </div>

      <div className="grid grid-cols-2 gap-2">
        {/* Option 1: Classic Image 2 Ornament */}
        <button
          type="button"
          onClick={() => handleSelectType('image2-classic')}
          className={`flex flex-col items-center justify-center p-2 rounded border transition text-center ${
            currentType === 'image2-classic'
              ? 'bg-white border-[#007bff] ring-1 ring-[#007bff] text-blue-900 font-bold'
              : 'bg-white/80 border-gray-200 text-gray-700 hover:bg-white'
          }`}
        >
          <div className="h-4 w-full flex items-center justify-center my-0.5">
            <TacteingLine type="image2-classic" width={110} height={10} />
          </div>
          <span className="text-[11px] mt-1">តាក់តែងបុរាណ (Classic)</span>
        </button>

        {/* Option 2: Symbol Cross -( + )- */}
        <button
          type="button"
          onClick={() => handleSelectType('symbol-cross')}
          className={`flex flex-col items-center justify-center p-2 rounded border transition text-center ${
            currentType === 'symbol-cross'
              ? 'bg-white border-[#007bff] ring-1 ring-[#007bff] text-blue-900 font-bold'
              : 'bg-white/80 border-gray-200 text-gray-700 hover:bg-white'
          }`}
        >
          <div className="h-4 w-full flex items-center justify-center my-0.5 font-mono text-[11px]">
            ───( + )───
          </div>
          <span className="text-[11px] mt-1">បន្ទាត់ -( + )-</span>
        </button>

        {/* Option 3: Symbol Flower -( ❁ )- */}
        <button
          type="button"
          onClick={() => handleSelectType('symbol-flower')}
          className={`flex flex-col items-center justify-center p-2 rounded border transition text-center ${
            currentType === 'symbol-flower'
              ? 'bg-white border-[#007bff] ring-1 ring-[#007bff] text-blue-900 font-bold'
              : 'bg-white/80 border-gray-200 text-gray-700 hover:bg-white'
          }`}
        >
          <div className="h-4 w-full flex items-center justify-center my-0.5 font-mono text-[11px]">
            ───( ❁ )───
          </div>
          <span className="text-[11px] mt-1">បន្ទាត់ -( ❁ )-</span>
        </button>

        {/* Option 4: Custom Upload */}
        <label
          className={`flex flex-col items-center justify-center p-2 rounded border cursor-pointer transition text-center relative ${
            currentType === 'custom-image' && customImage
              ? 'bg-white border-[#007bff] ring-1 ring-[#007bff] text-blue-900 font-bold'
              : 'bg-white/80 border-gray-200 text-gray-700 hover:bg-white'
          }`}
        >
          <input
            type="file"
            accept="image/*"
            onChange={handleFileUpload}
            className="hidden"
          />
          {currentType === 'custom-image' && customImage ? (
            <div className="h-4 w-full flex items-center justify-center my-0.5">
              <img src={customImage} alt="Custom" className="h-3.5 max-w-[100px] object-contain" />
            </div>
          ) : (
            <div className="flex items-center gap-1 text-[11px] my-0.5 text-blue-600 font-semibold">
              <Upload className="w-3 h-3" />
              <span>បញ្ចូលរូបភាព</span>
            </div>
          )}
          <span className="text-[11px] mt-1">រូបភាពផ្ទាល់ខ្លួន (Custom File)</span>
        </label>
      </div>
    </div>
  );
};
