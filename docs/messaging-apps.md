# Chat apps

Message your bots from Slack, Microsoft Teams, Telegram, WhatsApp, iMessage or SMS, and add them
to group chats there.

> [!NOTE]
> **Settings → Messaging** appears only when your server has at least one chat app set up. Which
> apps you see depends on the server: Slack, Microsoft Teams, Telegram, WhatsApp, iMessage and SMS,
> and Feishu (Lark) are supported. On a self-hosted server the operator configures them; see
> [Self-hosting](./self-host.md).

## Link a bot to a chat app

1. Open **Settings → Messaging**.
2. Pick the bot, then choose the app under **Connect a chat app**.
3. Send the code shown from that app within 10 minutes.

From then on, messages you send there go to that bot's chat, and its replies come back to you there.
The conversation is the same one you see in the Rakazo app.

A server that signs people in with Microsoft Entra can set `RAKAZO_TEAMS_AUTO_LINK=true`: the first
direct message from someone who has signed in to the server with the same Microsoft account links
Teams to their first bot that has no chat app yet, without a code, and the bot answers it. Anyone
else links with a code as usual.

The group intro and the prompt to join a group name the product as Rakazo. Set
`RAKAZO_PRODUCT_NAME` (up to 40 characters) to use another name.

In direct messages, a voice note the chat app delivers as an audio file (Telegram, for example)
is transcribed when a voice provider that supports transcription is connected in **Voice**
settings; without one, the bot receives the audio file.

To stop, choose **Unlink**.

## Slack

Choose **Connect** next to Slack to add Rakazo to your Slack workspace, then **Link DMs** so direct
messages reach your bot.

The bot uses its memory and notes only for your own messages: the Slack account that connected the
workspace or linked its DMs. It answers anyone else in the workspace, including guests and other
apps, the way it answers a group chat.

A [routine](./routines.md) can also start from a **Slack message**.

## Group chats

Add your bot to a group or channel in the chat app. The invite appears under **Requests** in
**Settings → Messaging**; choose **Approve** or **Decline**. Approved groups are listed under
**Conversations**, where you can **Leave** one at any time.

You can also answer an invite or a connection request by replying YES or NO to its message, within a
day and before anything else arrives from the line.

On Telegram, groups link personally, like on iMessage: each member who has linked the app gets
their own invite under **Requests**, and messages from approved linked members reach each approved
member's own bot. Members who have not linked Telegram only trigger a one-time intro note in the
group. Turn privacy mode off so the bot hears the whole group: in BotFather run `/setprivacy` and
choose **Disable**, or make the bot a group admin. With privacy mode on it only sees messages that
mention it. Telegram broadcast channels are not group chats; posts there never reach your bot.

Microsoft Teams group chats and channels link the same way. Install the bot's Teams app in a chat
or team and send it a message there. Unless your server's Teams app is set up to read every message,
the bot only sees messages that @mention it. In a channel, every thread belongs to one conversation
and the bot's replies start a new post in the channel. Files people share in group chats and channels
don't reach the bot; files in a direct chat with it do.

When someone in a group asks the bot to change something in one of your connected apps, you approve
it in the app first. A server can instead keep group chats read-only with
`RAKAZO_GROUP_RUNS_READ_ONLY=true`: the bot only reads connected apps there and asks people to
make changes from a direct chat.

Slack rooms added through a workspace are team conversations, the only ones with per-room
settings. Open a conversation's **Settings** to decide how the bot behaves there:

- **Listening**: **Listen** to everything, or **Mentions only**. Either way, the bot answers when
  mentioned directly.
- **Room guidance**: plain-language rules for the room, such as "Engage when someone asks about
  invoices. Ignore small talk."
- **Automated senders**: for each bot or integration that posts in the room, choose **Ignore**,
  **Group updates** into a periodic digest, **Always act**, or **Treat like a person**.
