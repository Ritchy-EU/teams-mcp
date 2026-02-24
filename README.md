# Teams MCP

A Model Context Protocol (MCP) server that provides seamless integration with Microsoft Graph APIs, enabling AI assistants to interact with Microsoft Teams, users, and organizational data.

## 📦 Installation

To use this MCP server in Cursor/Claude/VS Code, add the following configuration:

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@ritchy-eu/teams-mcp@latest"]
    }
  }
}
```

## 🚀 Features

### 🔐 Authentication
- OAuth 2.0 authentication flow with Microsoft Graph
- Secure token management and refresh
- Authentication status checking

### 👥 User Management
- Get current user information
- Search users by name or email
- Retrieve detailed user profiles
- Access organizational directory data

### 🏢 Microsoft Teams Integration
- **Teams Management**
  - List user's joined teams
  - Access team details and metadata
  
- **Channel Operations**
  - List channels within teams
  - Retrieve channel messages
  - Send messages to team channels
  - Support for message importance levels (normal, high, urgent)
  
- **Team Members**
  - List team members and their roles
  - Access member information

### 💬 Chat & Messaging
- **1:1 and Group Chats**
  - List user's chats
  - Create new 1:1 or group conversations
  - Retrieve chat message history with filtering and pagination
  - Send messages to existing chats
  - Edit previously sent chat messages
  - Soft delete chat messages

### ✏️ Message Management
- **Edit & Delete**
  - Update (edit) sent messages in chats and channels
  - Soft delete messages in chats and channels (marks as deleted without permanent removal)
  - Only message senders can update/delete their own messages
  - Support for Markdown formatting, mentions, and importance levels on edits

### 📎 Media & Attachments
- **Hosted Content**
  - Download hosted content (images, files) from chat and channel messages
  - Access inline images and attachments shared in conversations

### 🔍 Advanced Search & Discovery
- **Message Search**
  - Search across all Teams channels and chats using Microsoft Search API
  - Support for KQL (Keyword Query Language) syntax
  - Filter by sender, mentions, attachments, importance, and date ranges
  - Get recent messages with advanced filtering options
  - Find messages mentioning specific users

## Rich Message Formatting Support

The following tools now support rich message formatting in Teams channels and chats:
- `send_channel_message`
- `send_chat_message`
- `reply_to_channel_message`
- `update_channel_message`
- `update_chat_message`

### Format Options

You can specify the `format` parameter to control the message formatting:
- `text` (default): Plain text
- `markdown`: Markdown formatting (bold, italic, lists, links, code, etc.) - converted to sanitized HTML

When `format` is set to `markdown`, the message content is converted to HTML using a secure markdown parser and sanitized to remove potentially dangerous content before being sent to Teams.

If `format` is not specified, the message will be sent as plain text.

### Example Usage

```json
{
  "teamId": "...",
  "channelId": "...",
  "message": "**Bold text** and _italic text_\n\n- List item 1\n- List item 2\n\n[Link](https://example.com)",
  "format": "markdown"
}
```

```json
{
  "chatId": "...",
  "message": "Simple plain text message",
  "format": "text"
}
```

### Security Features

- **HTML Sanitization**: All markdown content is converted to HTML and sanitized to remove potentially dangerous elements (scripts, event handlers, etc.)
- **Allowed Tags**: Only safe HTML tags are permitted (p, strong, em, a, ul, ol, li, h1-h6, code, pre, etc.)
- **Safe Attributes**: Only safe attributes are allowed (href, target, src, alt, title, width, height)
- **XSS Prevention**: Content is automatically sanitized to prevent cross-site scripting attacks

### Supported Markdown Features

- **Text formatting**: Bold (`**text**`), italic (`_text_`), strikethrough (`~~text~~`)
- **Links**: `[text](url)` 
- **Lists**: Bulleted (`- item`) and numbered (`1. item`)
- **Code**: Inline `` `code` `` and blocks ``` ```code``` ```
- **Headings**: `# H1` through `###### H6`
- **Line breaks**: Automatic conversion of newlines to `<br>` tags
- **Blockquotes**: `> quoted text`
- **Tables**: GitHub-flavored markdown tables

## 📦 Installation

```bash
# Install dependencies
npm install

# Build the project
npm run build

# Set up authentication
npm run auth
```

## 🔧 Configuration

### Prerequisites
- Node.js 18+
- Microsoft 365 account with appropriate permissions
- Azure App Registration with Microsoft Graph permissions

### Required Microsoft Graph Permissions
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read basic user info
- `Team.ReadBasic.All` - Read team information
- `Channel.ReadBasic.All` - Read channel information
- `ChannelMessage.Read.All` - Read channel messages
- `ChannelMessage.Send` - Send channel messages
- `ChannelMessage.ReadWrite` - Edit and delete channel messages
- `Chat.Read` - Read chat messages
- `Chat.ReadWrite` - Create and manage chats (including edit/delete messages)
- `Mail.Read` - Required for Microsoft Search API
- `Calendars.Read` - Required for Microsoft Search API
- `Files.Read.All` - Required for Microsoft Search API
- `Sites.Read.All` - Required for Microsoft Search API

## 🛠️ Usage

### Starting the Server
```bash
# Development mode with hot reload
npm run dev

# Production mode
npm run build && node dist/index.js
```

### Available MCP Tools

#### Authentication
- `authenticate` - Initiate OAuth authentication flow
- `logout` - Clear authentication tokens
- `get_current_user` - Get authenticated user information

#### User Operations
- `search_users` - Search for users by name or email
- `get_user` - Get detailed user information by ID or email

#### Teams Operations
- `list_teams` - List user's joined teams
- `list_channels` - List channels in a specific team
- `get_channel_messages` - Retrieve messages from a team channel with pagination and filtering
- `send_channel_message` - Send a message to a team channel
- `update_channel_message` - Edit a previously sent channel message
- `delete_channel_message` - Soft delete a channel message (supports replies)
- `list_team_members` - List members of a specific team

#### Chat Operations
- `list_chats` - List user's chats (1:1 and group)
- `get_chat_messages` - Retrieve messages from a specific chat with pagination and filtering
- `send_chat_message` - Send a message to a chat
- `create_chat` - Create a new 1:1 or group chat
- `update_chat_message` - Edit a previously sent chat message
- `delete_chat_message` - Soft delete a chat message

#### Media Operations
- `download_message_hosted_content` - Download hosted content (images, files) from messages

#### Search Operations
- `search_messages` - Search across all Teams messages using KQL syntax
- `get_recent_messages` - Get recent messages with advanced filtering options
- `get_my_mentions` - Find messages mentioning the current user

## 📋 Examples

### Authentication

First, authenticate with Microsoft Graph:

```bash
npx @ritchy-eu/teams-mcp@latest authenticate
```

Check your authentication status:

```bash
npx @ritchy-eu/teams-mcp@latest check
```

Logout if needed:

```bash
npx @ritchy-eu/teams-mcp@latest logout
```

### Integrating with Cursor/Claude

This MCP server is designed to work with AI assistants like Claude/Cursor/VS Code through the Model Context Protocol. 

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@ritchy-eu/teams-mcp@latest"]
    }
  }
}
```

## 🔒 Security

- All authentication is handled through Microsoft's OAuth 2.0 flow
- **Refresh token support**: Access tokens are automatically renewed using cached refresh tokens, so you don't need to re-authenticate every hour
- Token cache is stored locally at `~/.teams-mcp-token-cache.json`
- No sensitive data is logged or exposed
- Follows Microsoft Graph API security best practices

## 📝 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run linting and formatting
5. Submit a pull request

## 📞 Support

For issues and questions:
- Check the existing GitHub issues
- Review Microsoft Graph API documentation
- Ensure proper authentication and permissions are configured