// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Content, Root } from 'chat'
import { BaseFormatConverter, isTextNode, parseMarkdown, stringifyMarkdown } from 'chat'

/**
 * Format converter for Microsoft Teams.
 *
 * Teams supports a subset of HTML for rich text and also accepts markdown
 * (`textFormat: "markdown"`). This converter uses markdown as the
 * interchange format since it is supported natively by the Bot Framework.
 *
 * Incoming Teams HTML (mentions, bold, italic, etc.) is normalised to
 * markdown before parsing into an mdast tree.
 *
 * Outgoing messages convert `@name` text nodes back to Teams `<at>name</at>`
 * mention pills so they render as proper mentions in the Teams client.
 */
export class TeamsFormatConverter extends BaseFormatConverter {
  toAst (platformText: string): Root {
    return parseMarkdown(teamsHtmlToMarkdown(platformText))
  }

  fromAst (ast: Root): string {
    return this.fromAstWithNodeConverter(ast, (node) => this.nodeToTeams(node))
  }

  private nodeToTeams (node: Content): string {
    // Convert plain-text @-mentions to Teams <at> HTML mention pills
    if (isTextNode(node) && node.value.startsWith('@')) {
      const name = node.value.slice(1)
      if (name.length > 0) {
        return `<at>${name}</at>`
      }
    }
    return this.defaultNodeToText(node, (n) => this.nodeToTeams(n))
  }
}

/**
 * Convert a Teams HTML-formatted string to markdown.
 * Teams uses a restricted set of HTML tags for rich messages.
 */
export function teamsHtmlToMarkdown (html: string): string {
  return html
    // Mentions: <at>Display Name</at> → @Display Name
    .replace(/<at[^>]*>(.*?)<\/at>/gi, '@$1')
    // Bold
    .replace(/<strong>(.*?)<\/strong>/gi, '**$1**')
    .replace(/<b>(.*?)<\/b>/gi, '**$1**')
    // Italic
    .replace(/<em>(.*?)<\/em>/gi, '_$1_')
    .replace(/<i>(.*?)<\/i>/gi, '_$1_')
    // Strikethrough
    .replace(/<s>(.*?)<\/s>/gi, '~~$1~~')
    .replace(/<del>(.*?)<\/del>/gi, '~~$1~~')
    // Inline code
    .replace(/<code>(.*?)<\/code>/gi, '`$1`')
    // Code blocks
    .replace(/<pre><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, '```\n$1\n```')
    // Links
    .replace(/<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)')
    // Line breaks
    .replace(/<br\s*\/?>/gi, '\n')
    // Block elements with newlines
    .replace(/<\/(p|div|h[1-6])>/gi, '\n')
    .replace(/<(p|div|h[1-6])[^>]*>/gi, '')
    // Strip any remaining HTML tags
    .replace(/<[^>]+>/g, '')
    // Decode common HTML entities
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .trim()
}

// Re-export for callers that relied on the old stringifyMarkdown-based fromAst
export { stringifyMarkdown }
