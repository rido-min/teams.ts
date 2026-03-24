// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type {
  ActionsElement,
  ButtonElement,
  CardChild,
  CardElement,
  DividerElement,
  FieldsElement,
  ImageElement,
  LinkButtonElement,
  RadioSelectElement,
  SectionElement,
  SelectElement,
  TableElement,
  TextElement,
} from 'chat'

// ── Adaptive Card schema types (v1.5) ────────────────────────────────────────

type AdaptiveCardBody = AdaptiveTextBlock | AdaptiveImage | AdaptiveContainer | AdaptiveColumnSet | AdaptiveFactSet | AdaptiveSeparator | AdaptiveTable

interface AdaptiveTextBlock {
  type: 'TextBlock'
  text: string
  weight?: 'bolder' | 'default'
  color?: 'default' | 'subtle'
  isSubtle?: boolean
  wrap: true
}

interface AdaptiveImage {
  type: 'Image'
  url: string
  altText?: string
}

interface AdaptiveSeparator {
  type: 'TextBlock'
  text: ' '
  separator: true
}

interface AdaptiveContainer {
  type: 'Container'
  items: AdaptiveCardBody[]
}

interface AdaptiveColumn {
  type: 'Column'
  width: string
  items: AdaptiveCardBody[]
}

interface AdaptiveColumnSet {
  type: 'ColumnSet'
  columns: AdaptiveColumn[]
}

interface AdaptiveFactSet {
  type: 'FactSet'
  facts: Array<{ title: string; value: string }>
}

interface AdaptiveTableCell {
  type: 'TableCell'
  items: AdaptiveCardBody[]
}

interface AdaptiveTableRow {
  type: 'TableRow'
  cells: AdaptiveTableCell[]
}

interface AdaptiveTable {
  type: 'Table'
  columns: Array<{ width: number }>
  rows: AdaptiveTableRow[]
  showGridLines: boolean
}

type AdaptiveAction =
  | { type: 'Action.Submit'; title: string; style?: string; data: Record<string, unknown> }
  | { type: 'Action.OpenUrl'; title: string; style?: string; url: string }
  | { type: 'Action.ShowCard'; title: string; card: AdaptiveCardObject }
  | AdaptiveChoiceSetInput

interface AdaptiveChoiceSetInput {
  type: 'Input.ChoiceSet'
  id: string
  label: string
  placeholder?: string
  isRequired?: boolean
  value?: string
  choices: Array<{ title: string; value: string }>
  style?: 'compact' | 'expanded'
}

interface AdaptiveCardObject {
  type: 'AdaptiveCard'
  version: '1.5'
  body: AdaptiveCardBody[]
  actions: AdaptiveAction[]
}

// ── Converter ────────────────────────────────────────────────────────────────

/**
 * Converts a Chat SDK {@link CardElement} to an Adaptive Card v1.5 JSON object
 * suitable for sending as a `application/vnd.microsoft.card.adaptive` attachment
 * in Microsoft Teams.
 */
export function cardToAdaptiveCard (card: CardElement): AdaptiveCardObject {
  const body: AdaptiveCardBody[] = []
  const actions: AdaptiveAction[] = []

  if (card.title) {
    body.push({ type: 'TextBlock', text: card.title, weight: 'bolder', wrap: true })
  }
  if (card.subtitle) {
    body.push({ type: 'TextBlock', text: card.subtitle, isSubtle: true, wrap: true })
  }
  if (card.imageUrl) {
    body.push({ type: 'Image', url: card.imageUrl })
  }

  for (const child of card.children) {
    const converted = cardChildToBody(child)
    if (converted === null) continue
    if (Array.isArray(converted)) {
      // ActionsElement produces actions, not body items
      for (const a of converted) actions.push(a)
    } else {
      body.push(converted)
    }
  }

  return { type: 'AdaptiveCard', version: '1.5', body, actions }
}

// Returns null to skip, AdaptiveCardBody for body items, or AdaptiveAction[] for actions
function cardChildToBody (child: CardChild): AdaptiveCardBody | AdaptiveAction[] | null {
  switch (child.type) {
    case 'text':
      return textElementToBody(child)
    case 'image':
      return imageElementToBody(child)
    case 'divider':
      return dividerToBody(child)
    case 'section':
      return sectionToBody(child)
    case 'fields':
      return fieldsToBody(child)
    case 'actions':
      return actionsToAdaptiveActions(child)
    case 'link':
      return { type: 'TextBlock', text: `[${child.label}](${child.url})`, wrap: true }
    case 'table':
      return tableToBody(child)
    default:
      return null
  }
}

function textElementToBody (el: TextElement): AdaptiveTextBlock {
  const block: AdaptiveTextBlock = { type: 'TextBlock', text: el.content, wrap: true }
  if (el.style === 'bold') block.weight = 'bolder'
  if (el.style === 'muted') block.isSubtle = true
  return block
}

function imageElementToBody (el: ImageElement): AdaptiveImage {
  return { type: 'Image', url: el.url, altText: el.alt }
}

function dividerToBody (_el: DividerElement): AdaptiveSeparator {
  return { type: 'TextBlock', text: ' ', separator: true }
}

function sectionToBody (el: SectionElement): AdaptiveContainer {
  const items: AdaptiveCardBody[] = []
  for (const child of el.children) {
    const converted = cardChildToBody(child)
    if (converted === null || Array.isArray(converted)) continue
    items.push(converted)
  }
  return { type: 'Container', items }
}

function fieldsToBody (el: FieldsElement): AdaptiveFactSet {
  return {
    type: 'FactSet',
    facts: el.children.map((f) => ({ title: f.label, value: f.value })),
  }
}

function actionsToAdaptiveActions (el: ActionsElement): AdaptiveAction[] {
  const actions: AdaptiveAction[] = []
  for (const child of el.children) {
    switch (child.type) {
      case 'button':
        actions.push(buttonToAction(child))
        break
      case 'link-button':
        actions.push(linkButtonToAction(child))
        break
      case 'select':
        actions.push(selectToChoiceSet(child))
        break
      case 'radio_select':
        actions.push(radioSelectToChoiceSet(child))
        break
    }
  }
  return actions
}

function buttonToAction (el: ButtonElement): AdaptiveAction {
  const action: Extract<AdaptiveAction, { type: 'Action.Submit' }> = {
    type: 'Action.Submit',
    title: el.label,
    data: { actionId: el.id, value: el.value },
  }
  if (el.style === 'primary') action.style = 'positive'
  if (el.style === 'danger') action.style = 'destructive'
  return action
}

function linkButtonToAction (el: LinkButtonElement): AdaptiveAction {
  const action: Extract<AdaptiveAction, { type: 'Action.OpenUrl' }> = {
    type: 'Action.OpenUrl',
    title: el.label,
    url: el.url,
  }
  if (el.style === 'primary') action.style = 'positive'
  if (el.style === 'danger') action.style = 'destructive'
  return action
}

function selectToChoiceSet (el: SelectElement): AdaptiveChoiceSetInput {
  return {
    type: 'Input.ChoiceSet',
    id: el.id,
    label: el.label,
    placeholder: el.placeholder,
    isRequired: !el.optional,
    value: el.initialOption,
    choices: el.options.map((o) => ({ title: o.label, value: o.value })),
    style: 'compact',
  }
}

function radioSelectToChoiceSet (el: RadioSelectElement): AdaptiveChoiceSetInput {
  return {
    type: 'Input.ChoiceSet',
    id: el.id,
    label: el.label,
    isRequired: !el.optional,
    value: el.initialOption,
    choices: el.options.map((o) => ({ title: o.label, value: o.value })),
    style: 'expanded',
  }
}

function tableToBody (el: TableElement): AdaptiveTable {
  const colCount = el.headers.length

  const headerRow: AdaptiveTableRow = {
    type: 'TableRow',
    cells: el.headers.map((h) => ({
      type: 'TableCell',
      items: [{ type: 'TextBlock', text: h, weight: 'bolder', wrap: true }],
    })),
  }

  const dataRows: AdaptiveTableRow[] = el.rows.map((row) => ({
    type: 'TableRow',
    cells: row.map((cell) => ({
      type: 'TableCell',
      items: [{ type: 'TextBlock', text: cell, wrap: true }],
    })),
  }))

  return {
    type: 'Table',
    columns: Array.from({ length: colCount }, () => ({ width: 1 })),
    rows: [headerRow, ...dataRows],
    showGridLines: true,
  }
}
