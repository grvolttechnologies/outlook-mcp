export async function authenticateTool(authManager) {
  const result = await authManager.authenticate();
  
  if (result.success) {
    return {
      content: [
        {
          type: 'text',
          text: `Successfully authenticated as ${result.user.displayName} (${result.user.mail})`,
        },
      ],
    };
  } else {
    return {
      error: {
        code: 'AUTH_FAILED',
        message: result.error,
      },
    };
  }
}

export async function listEmailsTool(authManager, args) {
  const { folder = 'inbox', limit = 10, filter } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'subject,from,receivedDateTime,bodyPreview,isRead',
      top: limit,
      orderby: 'receivedDateTime desc',
    };

    if (filter) {
      options.filter = filter;
    }

    const result = await graphApiClient.makeRequest(`/me/mailFolders/${folder}/messages`, options);

    const emails = result.value.map(email => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.emailAddress?.address || 'Unknown',
      fromName: email.from?.emailAddress?.name || 'Unknown',
      receivedDateTime: email.receivedDateTime,
      preview: email.bodyPreview,
      isRead: email.isRead,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ emails, count: emails.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list emails: ${error.message}`);
  }
}

export async function sendEmailTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [] } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const message = {
      subject,
      body: {
        contentType: bodyType === 'html' ? 'HTML' : 'Text',
        content: body,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    if (cc.length > 0) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    await graphApiClient.postWithRetry('/me/sendMail', {
      message,
      saveToSentItems: true,
    });

    return {
      content: [
        {
          type: 'text',
          text: `Email sent successfully to ${to.join(', ')}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to send email: ${error.message}`);
  }
}

export async function listEventsTool(authManager, args) {
  const { startDateTime, endDateTime, limit = 10, calendar } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendar ? `/me/calendars/${calendar}/events` : '/me/events';
    const options = {
      select: 'subject,start,end,location,attendees,bodyPreview',
      top: limit,
      orderby: 'start/dateTime',
    };

    if (startDateTime && endDateTime) {
      options.filter = `start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`;
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
      preview: event.bodyPreview,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ events, count: events.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list events: ${error.message}`);
  }
}

export async function createEventTool(authManager, args) {
  const { subject, start, end, body = '', location = '', attendees = [] } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const event = {
      subject,
      start,
      end,
      body: {
        contentType: 'Text',
        content: body,
      },
    };

    if (location) {
      event.location = {
        displayName: location,
      };
    }

    if (attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    return {
      content: [
        {
          type: 'text',
          text: `Event "${subject}" created successfully. Event ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create event: ${error.message}`);
  }
}