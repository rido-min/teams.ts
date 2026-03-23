// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it, mock } from 'node:test'
import assert from 'node:assert/strict'
import { ConversationClient } from './conversation-client.js'

// Minimal test: constructor creates a client and methods exist
describe('ConversationClient', () => {
  it('can be instantiated without a token provider', () => {
    const client = new ConversationClient()
    assert.ok(client)
  })

  it('can be instantiated with a token provider', () => {
    const client = new ConversationClient(async () => 'test-token')
    assert.ok(client)
  })

  it('exposes expected methods', () => {
    const client = new ConversationClient()
    assert.equal(typeof client.sendActivityAsync, 'function')
    assert.equal(typeof client.updateActivityAsync, 'function')
    assert.equal(typeof client.deleteActivityAsync, 'function')
    assert.equal(typeof client.getConversationMembersAsync, 'function')
    assert.equal(typeof client.getConversationMemberAsync, 'function')
    assert.equal(typeof client.getConversationPagedMembersAsync, 'function')
    assert.equal(typeof client.deleteConversationMemberAsync, 'function')
    assert.equal(typeof client.createConversationAsync, 'function')
    assert.equal(typeof client.getConversationsAsync, 'function')
    assert.equal(typeof client.sendConversationHistoryAsync, 'function')
    assert.equal(typeof client.addReactionAsync, 'function')
    assert.equal(typeof client.deleteReactionAsync, 'function')
    assert.equal(typeof client.uploadAttachmentAsync, 'function')
  })

  it('mock is available', () => {
    assert.equal(typeof mock, 'object')
  })
})
