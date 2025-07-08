// Helper functions for common Graph API operations

export const graphHelpers = {
  // Email helpers
  email: {
    buildMessageObject(to, subject, body, options = {}) {
      const message = {
        subject,
        body: {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: body,
        },
        toRecipients: Array.isArray(to) ? 
          to.map(email => ({ emailAddress: { address: email } })) :
          [{ emailAddress: { address: to } }],
      };

      if (options.cc) {
        message.ccRecipients = options.cc.map(email => ({
          emailAddress: { address: email },
        }));
      }

      if (options.bcc) {
        message.bccRecipients = options.bcc.map(email => ({
          emailAddress: { address: email },
        }));
      }

      if (options.importance) {
        message.importance = options.importance; // low, normal, high
      }

      if (options.attachments) {
        message.attachments = options.attachments;
      }

      return message;
    },

    buildReplyObject(body, options = {}) {
      const reply = {
        comment: body,
      };

      if (options.replyAll) {
        reply.message = {};
        if (options.cc) {
          reply.message.ccRecipients = options.cc.map(email => ({
            emailAddress: { address: email },
          }));
        }
      }

      return reply;
    },

    parseEmailAddress(emailObject) {
      if (typeof emailObject === 'string') return emailObject;
      return emailObject?.emailAddress?.address || 'unknown';
    },

    parseEmailName(emailObject) {
      if (typeof emailObject === 'string') return null;
      return emailObject?.emailAddress?.name || null;
    },
  },

  // Calendar helpers
  calendar: {
    buildEventObject(subject, start, end, options = {}) {
      const event = {
        subject,
        start: {
          dateTime: start.dateTime || start,
          timeZone: start.timeZone || 'UTC',
        },
        end: {
          dateTime: end.dateTime || end,
          timeZone: end.timeZone || 'UTC',
        },
      };

      if (options.body) {
        event.body = {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: options.body,
        };
      }

      if (options.location) {
        event.location = {
          displayName: options.location,
        };
      }

      if (options.attendees) {
        event.attendees = options.attendees.map(email => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      if (options.isAllDay) {
        event.isAllDay = true;
      }

      if (options.recurrence) {
        event.recurrence = options.recurrence;
      }

      if (options.isOnlineMeeting) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = options.onlineMeetingProvider || 'teamsForBusiness';
      }

      return event;
    },

    buildRecurrencePattern(pattern, range) {
      const recurrence = {
        pattern: {
          type: pattern.type, // daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly
          interval: pattern.interval || 1,
        },
        range: {
          type: range.type, // endDate, noEnd, numbered
          startDate: range.startDate,
        },
      };

      if (pattern.daysOfWeek) {
        recurrence.pattern.daysOfWeek = pattern.daysOfWeek;
      }

      if (pattern.dayOfMonth) {
        recurrence.pattern.dayOfMonth = pattern.dayOfMonth;
      }

      if (range.type === 'endDate') {
        recurrence.range.endDate = range.endDate;
      } else if (range.type === 'numbered') {
        recurrence.range.numberOfOccurrences = range.numberOfOccurrences;
      }

      return recurrence;
    },

    parseDateTimeWithZone(dateTime, timeZone = 'UTC') {
      return {
        dateTime: dateTime,
        timeZone: timeZone,
      };
    },
  },

  // Contact helpers
  contact: {
    buildContactObject(givenName, surname, options = {}) {
      const contact = {
        givenName,
        surname,
      };

      if (options.displayName) {
        contact.displayName = options.displayName;
      } else {
        contact.displayName = `${givenName} ${surname}`;
      }

      if (options.emailAddresses) {
        contact.emailAddresses = options.emailAddresses.map(email => ({
          address: email.address || email,
          name: email.name || contact.displayName,
        }));
      }

      if (options.businessPhones) {
        contact.businessPhones = Array.isArray(options.businessPhones) 
          ? options.businessPhones 
          : [options.businessPhones];
      }

      if (options.mobilePhone) {
        contact.mobilePhone = options.mobilePhone;
      }

      if (options.jobTitle) {
        contact.jobTitle = options.jobTitle;
      }

      if (options.companyName) {
        contact.companyName = options.companyName;
      }

      if (options.department) {
        contact.department = options.department;
      }

      if (options.businessAddress) {
        contact.businessAddress = options.businessAddress;
      }

      return contact;
    },
  },

  // Task helpers
  task: {
    buildTaskObject(title, options = {}) {
      const task = {
        title,
        status: options.status || 'notStarted', // notStarted, inProgress, completed, waitingOnOthers, deferred
      };

      if (options.body) {
        task.body = {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: options.body,
        };
      }

      if (options.dueDateTime) {
        task.dueDateTime = {
          dateTime: options.dueDateTime,
          timeZone: options.timeZone || 'UTC',
        };
      }

      if (options.startDateTime) {
        task.startDateTime = {
          dateTime: options.startDateTime,
          timeZone: options.timeZone || 'UTC',
        };
      }

      if (options.importance) {
        task.importance = options.importance; // low, normal, high
      }

      if (options.recurrence) {
        task.recurrence = options.recurrence;
      }

      if (options.categories) {
        task.categories = options.categories;
      }

      return task;
    },
  },

  // General helpers
  general: {
    buildODataFilter(filters) {
      if (!filters || Object.keys(filters).length === 0) return null;

      const filterStrings = [];

      for (const [key, value] of Object.entries(filters)) {
        if (value === null || value === undefined) continue;

        if (typeof value === 'string') {
          filterStrings.push(`${key} eq '${value}'`);
        } else if (typeof value === 'boolean') {
          filterStrings.push(`${key} eq ${value}`);
        } else if (value instanceof Date) {
          filterStrings.push(`${key} eq ${value.toISOString()}`);
        } else if (typeof value === 'object') {
          // Handle complex filters like { $gt: date }
          for (const [operator, val] of Object.entries(value)) {
            switch (operator) {
              case '$gt':
                filterStrings.push(`${key} gt ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$gte':
                filterStrings.push(`${key} ge ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$lt':
                filterStrings.push(`${key} lt ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$lte':
                filterStrings.push(`${key} le ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$ne':
                filterStrings.push(`${key} ne '${val}'`);
                break;
              case '$contains':
                filterStrings.push(`contains(${key}, '${val}')`);
                break;
              case '$startswith':
                filterStrings.push(`startswith(${key}, '${val}')`);
                break;
            }
          }
        }
      }

      return filterStrings.join(' and ');
    },

    parseGraphError(error) {
      if (error.body?.error) {
        return {
          code: error.body.error.code,
          message: error.body.error.message,
          innerError: error.body.error.innerError,
        };
      }
      return {
        code: 'Unknown',
        message: error.message || 'An unknown error occurred',
      };
    },

    // Format file size for display
    formatFileSize(bytes) {
      const sizes = ['Bytes', 'KB', 'MB', 'GB'];
      if (bytes === 0) return '0 Bytes';
      const i = Math.floor(Math.log(bytes) / Math.log(1024));
      return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
    },
  },
};