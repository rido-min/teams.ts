// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { BaseFormatConverter, parseMarkdown, stringifyMarkdown } from 'chat'
import type { Root } from 'chat'

/**
 * Format converter for Microsoft Teams.
 *
 * Teams supports a subset of HTML for rich text and also accepts markdown
 * (`textFormat: "markdown"`). This converter uses markdown as the
 * interchange format since it is supported natively by the Bot Framework.
 *
 * Incoming Teams HTML (mentions, bold, italic, etc.) is normalised to
 * markdown before parsing into an mdast tree.
 */
export class TeamsFormatConverter extends BaseFormatConverter {
  toAst (platformText: string): Root {
    return parseMarkdown(teamsHtmlToMarkdown(platformText))
  }

  fromAst (ast: Root): string {
    return stringifyMarkdown(ast)
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
