import { expect } from "bun:test";
import { visual, withVisualPage } from "./harness.js";
import { serveDocx } from "./docx-harness.js";
import { inheritedTableSections } from "./docx-table-style-harness.js";
import type { DocumentSection } from "../src/core/model.js";

visual("DOCX table style ancestry paints inherited borders and padding in flow and layout", async () => {
  const base = await inheritedTableSections();
  const overridden = await inheritedTableSections(
    '<w:tblPr><w:tblBorders><w:top w:val="nil"/></w:tblBorders>' +
    '<w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>',
    '<w:tblBorders><w:right w:val="none"/></w:tblBorders>' +
    '<w:tblCellMar><w:right w:w="0" w:type="dxa"/></w:tblCellMar>',
    '<w:tblPr><w:tblBorders><w:right w:val="double" w:sz="16" w:color="654321"/></w:tblBorders>' +
    '<w:tblCellMar><w:right w:w="720" w:type="dxa"/></w:tblCellMar></w:tblPr>',
  );
  // Render the real table by itself; unrelated float/media pages need not load.
  const onlyTable = (sections: DocumentSection[]): DocumentSection[] => {
    const table = sections.flatMap(section => section.root).filter(node => node.type === "table").at(-1)!;
    const section = sections.find(section => section.root.includes(table))!;
    return [{ properties: section.properties, root: [table] }];
  };
  const { server } = await serveDocx(`${import.meta.dir}/../../../fixtures/docx-coverage.docx`, `<!doctype html>
    <main id="host" style="width:960px"></main><script type="module">
    import {render} from '/docx.mjs';
    window.paint = async (model, mode) => {
      window.handle?.destroy();
      window.handle = await render({document:model,target:document.querySelector('#host'),pagination:{mode},zoom:1});
      const cells = [...document.querySelectorAll('[data-aared-element="table-cell"]')];
      const read = cell => {const s=getComputedStyle(cell);return {
        top:s.borderTopStyle,right:s.borderRightStyle,bottom:s.borderBottomStyle,left:s.borderLeftStyle,
        color:s.borderTopColor,paddingTop:parseFloat(s.paddingTop),paddingLeft:parseFloat(s.paddingLeft),paddingRight:parseFloat(s.paddingRight)};};
      return {first:read(cells[0]),last:read(cells[2]),body:read(cells[3]),count:cells.length};
    };</script>`);
  await withVisualPage({ server, width: 1000, height: 800 }, async (view) => {
    type Output = { first: {top:string;left:string;right:string;color:string;paddingTop:number;paddingLeft:number;paddingRight:number};
      last:{left:string;right:string;paddingLeft:number;paddingRight:number};body:{top:string;color:string};count:number };
    for (const mode of ["flow", "layout"]) {
      const output = await view.evaluate(`paint(${JSON.stringify(onlyTable(base))},${JSON.stringify(mode)})`) as Output;
      expect(output.count).toBe(17); // The real vertical merge remains intact.
      // The native table is RTL: logical start/end borders mirror, but its
      // left/right margins stay physical. Border centering adds <1px slack.
      expect(output.first).toMatchObject({ top: "solid", right: "solid", color: "rgb(18, 52, 86)" });
      expect(output.first.paddingTop).toBeCloseTo(9.6, 0);
      expect(Math.abs(output.first.paddingLeft - 38.4)).toBeLessThan(1);
      expect(Math.abs(output.last.paddingRight - 19.2)).toBeLessThan(1);
      expect(output.body).toMatchObject({ top: "solid", color: "rgb(18, 52, 86)" });
      const direct = await view.evaluate(`paint(${JSON.stringify(onlyTable(overridden))},${JSON.stringify(mode)})`) as Output;
      expect(direct.first.top).toBe("none");
      expect(direct.first.right).toBe("solid");
      expect(direct.first.paddingRight).toBe(0);
      expect(direct.last.left).toBe("none");
      expect(direct.last.paddingLeft).toBe(0);
      expect(direct.body).toMatchObject({ top: "solid", color: "rgb(18, 52, 86)" });
    }
    await view.evaluate("handle.destroy()");
  });
}, 30000);
