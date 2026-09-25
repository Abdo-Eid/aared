/**
 * Light syntax theme for the cream code cards. It follows GitHub light's
 * structure (red keywords, purple functions, blue constants, navy strings),
 * which reads well and matches GitHub dark, with three changes: plain text
 * uses the site's ink colour, HTML/JSX tag names are warm brown instead of
 * green, and red, purple and the comment grey are a shade darker so every
 * token holds 4.5:1 on the deeper parchment code blocks. Dark mode keeps
 * GitHub dark unchanged.
 */
const rule = (scope: string[], foreground: string, fontStyle?: string) => ({
  scope,
  settings: fontStyle ? { foreground, fontStyle } : { foreground },
});

export const aaredLight = {
  name: "aared-light",
  type: "light" as const,
  colors: { "editor.foreground": "#3c1215", "editor.background": "#f4ede1" },
  tokenColors: [
    rule(["comment", "punctuation.definition.comment"], "#6a5e57", "italic"),
    rule(
      [
        "keyword",
        "keyword.control",
        "keyword.operator.new",
        "keyword.operator.expression",
        "storage",
        "storage.type",
        "storage.modifier",
        "variable.language",
      ],
      "#bb1e2b",
    ),
    rule(["entity.name.tag", "support.class.component"], "#953800"),
    rule(
      [
        "entity.name.function",
        "support.function",
        "meta.function-call entity.name.function",
        "entity.other.attribute-name",
        "entity.name.type",
        "entity.name.class",
        "support.class",
        "support.type",
      ],
      "#7342cc",
    ),
    rule(
      [
        "constant",
        "constant.numeric",
        "constant.language",
        "support.constant",
        "constant.other",
        "meta.object-literal.key",
        "variable.other.property",
        "support.type.property-name",
        "meta.property-name",
      ],
      "#0550ae",
    ),
    rule(["string", "punctuation.definition.string", "string.quoted"], "#0a3069"),
    rule(["punctuation.definition.tag"], "#6a5e57"),
    rule(["variable", "variable.other", "variable.parameter", "meta.embedded", "punctuation", "meta.brace"], "#3c1215"),
  ],
};

/**
 * The one theme pair every code surface uses (markdown fences, `CodeBlock`,
 * the home page's light table). Rendered with `defaultColor: false`, so each
 * token carries `--shiki-light` / `--shiki-dark` and the site's `data-theme`
 * picks which applies — see the code rules in `DocsLayout` and `LightTable`.
 */
export const codeThemes = { light: aaredLight, dark: "github-dark" } as const;
