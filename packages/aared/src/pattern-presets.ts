/**
 * DrawingML's 54 preset patterns (`ST_PresetPatternVal`), then SpreadsheetML's
 * 17 cell patterns (`ST_PatternType`, below), as 8×8 bitmaps: 16
 * hex digits, one byte per row top to bottom, most significant bit leftmost,
 * a set bit painting the foreground colour. The spec names the presets but
 * gives no bitmaps, so these were read from PowerPoint itself: every preset
 * drawn black on white through its COM object model and exported at 1280 px
 * (96 px per inch) and 2560 px. The two exports agree once the second is
 * halved, every preset repeats within 8 px at 1280, and shapes offset off the
 * 8 px grid show the same bitmap shifted by the offset — so a cell is one
 * CSS pixel (the 8×8 tile is 1/12 inch), patterns scale with the slide, and
 * they are anchored to the slide's origin (Notes §28). `cross` and
 * `lgGrid` really are the same bitmap.
 */
export const PATTERN_PRESETS: Readonly<Record<string, string>> = {
  pct5: "0040000000040000",
  pct10: "0040000400400004",
  pct20: "0044001100440011",
  pct25: "1144114411441144",
  pct30: "8855225588552255",
  pct40: "15aa55aa51aa55aa",
  pct50: "aa55aa55aa55aa55",
  pct60: "aa77aaddaa77aadd",
  pct70: "eebbeebbeebbeebb",
  pct75: "ffbbffeeffbbffee",
  pct80: "ffeffffeffeffffe",
  pct90: "7ffffffff7ffffff",
  horz: "00ff000000000000",
  vert: "4040404040404040",
  ltHorz: "00ff000000ff0000",
  ltVert: "4444444444444444",
  dkHorz: "00ffff0000ffff00",
  dkVert: "6666666666666666",
  narHorz: "00ff00ff00ff00ff",
  narVert: "aaaaaaaaaaaaaaaa",
  dashHorz: "0078000000870000",
  dashVert: "0440404040040404",
  cross: "40ff404040404040",
  dnDiag: "8040201008040201",
  upDiag: "4080010204081020",
  ltDnDiag: "8844221188442211",
  ltUpDiag: "4488112244881122",
  dkDnDiag: "cc663399cc663399",
  dkUpDiag: "cc993366cc993366",
  wdDnDiag: "c1e070381c0e0783",
  wdUpDiag: "e0c183070e1c3870",
  dashDnDiag: "0000004422118800",
  dashUpDiag: "0000008811224400",
  diagCross: "8181422418182442",
  smCheck: "cccc3333cccc3333",
  lgCheck: "0ff0f0f0f00f0f0f",
  smGrid: "88ff888888ff8888",
  lgGrid: "40ff404040404040",
  dotGrid: "0055004000400040",
  smConfetti: "0240042001088010",
  lgConfetti: "c6d818818d6c6006",
  horzBrick: "08ff808080ff0808",
  diagBrick: "c0800102040c1221",
  solidDmnd: "00081c3e7f3e1c08",
  openDmnd: "8041221408142241",
  dotDmnd: "0080002200080022",
  plaid: "f0aa55aa55f0f0f0",
  sphere: "7cbbc4c7c7bb4c7c",
  weave: "a8442a11a2440a11",
  divot: "4000080408004080",
  shingle: "01038448300c0201",
  wave: "c0001825c0001825",
  trellis: "99ff66ff99ff66ff",
  zigZag: "0cc021120cc02112",
  // SpreadsheetML's 17 cell patterns (`ST_PatternType`), read from Excel's own
  // screen bitmaps (probe xlsx-cell-fills): each pattern black on white in a
  // large cell, copied with `Range.CopyPicture` at 100%, its repeating tile
  // taken from the cell's top-left. Every one repeats within 8 x 4 px, so
  // the tile is a CSS pixel per cell like DrawingML's. Excel anchors the tiles
  // to each cell (cells at odd pixel offsets show the same phase), not to the
  // sheet. darkGray, mediumGray and lightTrellis equal pct70, pct50 and pct30.
  darkGray: "eebbeebbeebbeebb",
  mediumGray: "aa55aa55aa55aa55",
  lightGray: "8822882288228822",
  gray125: "8800220088002200",
  gray0625: "8000080080000800",
  darkHorizontal: "ffff0000ffff0000",
  darkVertical: "cccccccccccccccc",
  darkDown: "cc663399cc663399",
  darkUp: "3366cc993366cc99",
  darkGrid: "9966669999666699",
  darkTrellis: "ff66ff99ff66ff99",
  lightHorizontal: "ff000000ff000000",
  lightVertical: "8888888888888888",
  lightDown: "8844221188442211",
  lightUp: "1122448811224488",
  lightGrid: "ff888888ff888888",
  lightTrellis: "8855225588552255",
};
