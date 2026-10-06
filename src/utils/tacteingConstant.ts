// Exact 100% reproduction of the Tacteing in IMG_8066.jpeg
// Structure:
// [Tapered Left Line] --- [Left Horizontal Eye with Dot] --- [Left Curved Bracket (] --- [Center 8-Ray Asterisk Rosette ❊] --- [Right Curved Bracket )] --- [Right Horizontal Eye with Dot] --- [Tapered Right Line]

export const KHMER_CLASSICAL_TACTEING_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 28" width="320" height="28" style="display:inline-block; vertical-align:middle;">
  <g fill="#000000" stroke="none" shape-rendering="geometricPrecision">
    <!-- 1. Left Horizontal Line with smooth taper at the outer tip -->
    <path d="M 20 14 L 128 12.8 L 128 15.2 Z" />

    <!-- 2. Left Eye / Leaf Shape (outer almond/ellipse outline with hollow interior) -->
    <path d="M 128 14 C 132 10.2, 142 10.2, 148 14 C 142 17.8, 132 17.8, 128 14 Z M 131 14 C 134 16.2, 140 16.2, 145 14 C 140 11.8, 134 11.8, 131 14 Z" fill-rule="evenodd" />

    <!-- Left Pupil / Dot inside the eye -->
    <circle cx="138" cy="14" r="2.2" />

    <!-- 3. Left Curved Bracket '(' -->
    <path d="M 152 7 C 148.5 11, 148.5 17, 152 21 C 153.2 21, 153.2 20.2, 152.2 19.4 C 149.8 16.2, 149.8 11.8, 152.2 8.6 C 153.2 7.8, 153.2 7, 152 7 Z" />

    <!-- 4. Center 8-Ray Asterisk Rosette Flower -->
    <!-- Center core dot -->
    <circle cx="160" cy="14" r="2.3" />

    <!-- 8 Radiating Teardrop / Oval Petals -->
    <!-- North (Top) -->
    <path d="M 160 10.2 C 158.5 8.2, 158.5 4.5, 160 3.2 C 161.5 4.5, 161.5 8.2, 160 10.2 Z" />
    <!-- South (Bottom) -->
    <path d="M 160 17.8 C 158.5 19.8, 158.5 23.5, 160 24.8 C 161.5 23.5, 161.5 19.8, 160 17.8 Z" />
    <!-- West (Left) -->
    <path d="M 156.2 14 C 154.2 12.5, 150.5 12.5, 149.2 14 C 150.5 15.5, 154.2 15.5, 156.2 14 Z" />
    <!-- East (Right) -->
    <path d="M 163.8 14 C 165.8 12.5, 169.5 12.5, 170.8 14 C 169.5 15.5, 165.8 15.5, 163.8 14 Z" />

    <!-- North-West (Top-Left) -->
    <path d="M 157.3 11.3 C 155 9.2, 152.2 6.5, 153.2 5.5 C 154.2 4.5, 157 7.2, 159.2 9.5 Z" />
    <!-- North-East (Top-Right) -->
    <path d="M 162.7 11.3 C 165 9.2, 167.8 6.5, 166.8 5.5 C 165.8 4.5, 163 7.2, 160.8 9.5 Z" />
    <!-- South-West (Bottom-Left) -->
    <path d="M 157.3 16.7 C 155 18.8, 152.2 21.5, 153.2 22.5 C 154.2 23.5, 157 20.8, 159.2 18.5 Z" />
    <!-- South-East (Bottom-Right) -->
    <path d="M 162.7 16.7 C 165 18.8, 167.8 21.5, 166.8 22.5 C 165.8 23.5, 163 20.8, 160.8 18.5 Z" />

    <!-- 5. Right Curved Bracket ')' -->
    <path d="M 168 7 C 171.5 11, 171.5 17, 168 21 C 166.8 21, 166.8 20.2, 167.8 19.4 C 170.2 16.2, 170.2 11.8, 167.8 8.6 C 166.8 7.8, 166.8 7, 168 7 Z" />

    <!-- 6. Right Eye / Leaf Shape (outer almond/ellipse outline with hollow interior) -->
    <path d="M 172 14 C 178 10.2, 188 10.2, 192 14 C 188 17.8, 178 17.8, 172 14 Z M 175 14 C 180 16.2, 186 16.2, 189 14 C 186 11.8, 180 11.8, 175 14 Z" fill-rule="evenodd" />

    <!-- Right Pupil / Dot inside the eye -->
    <circle cx="182" cy="14" r="2.2" />

    <!-- 7. Right Horizontal Line with smooth taper at the outer tip -->
    <path d="M 300 14 L 192 12.8 L 192 15.2 Z" />
  </g>
</svg>`;

export const DEFAULT_TACTEING_BASE64_DATA =
  'data:image/svg+xml;utf8,' + encodeURIComponent(KHMER_CLASSICAL_TACTEING_SVG);
