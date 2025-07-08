import { Client } from '@microsoft/microsoft-graph-client';
import { authConfig } from '../auth/config.js';

export class GraphApiClient {
  constructor(authManager) {
    this.authManager = authManager;
    this.client = null;
    this.requestCount = 0;
    this.requestWindow = [];
    this.maxConcurrentRequests = 4; // Per mailbox limit from Graph API
    this.activeRequests = 0;
  }

  async initialize() {
    if (this.client) return this.client;

    const authProvider = {
      getAccessToken: async () => {
        const tokenManager = this.authManager.tokenManager;
        try {
          return await tokenManager.getAccessToken();
        } catch (error) {
          if (error.message.includes('needs refresh')) {
            await this.authManager.refreshAccessToken();
            return await tokenManager.getAccessToken();
          }
          throw error;
        }
      },
    };

    this.client = Client.init({
      authProvider: (done) => {
        authProvider.getAccessToken()
          .then(token => done(null, token))
          .catch(error => done(error, null));
      },
      defaultVersion: 'v1.0',
      debugLogging: process.env.NODE_ENV === 'development',
    });

    this.setupMiddleware();
    return this.client;
  }

  setupMiddleware() {
    // Rate limiting middleware
    this.client.middleware.push({
      execute: async (context) => {
        await this.enforceRateLimit();
        return await context.next();
      },
      setNext: (next) => {},
    });

    // Retry middleware with exponential backoff
    this.client.middleware.push({
      execute: async (context) => {
        return await this.retryWithBackoff(context);
      },
      setNext: (next) => {},
    });

    // Correlation ID middleware
    this.client.middleware.push({
      execute: async (context) => {
        context.request.headers['client-request-id'] = this.generateCorrelationId();
        return await context.next();
      },
      setNext: (next) => {},
    });
  }

  async enforceRateLimit() {
    // Remove requests older than 1 minute
    const oneMinuteAgo = Date.now() - 60000;
    this.requestWindow = this.requestWindow.filter(time => time > oneMinuteAgo);

    // Wait if we're at the concurrent request limit
    while (this.activeRequests >= this.maxConcurrentRequests) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    this.activeRequests++;
    this.requestWindow.push(Date.now());
  }

  async retryWithBackoff(context) {
    const maxRetries = authConfig.retry.maxAttempts;
    let retryCount = 0;
    let delay = authConfig.retry.initialDelay;

    while (retryCount < maxRetries) {
      try {
        const response = await context.next();
        this.activeRequests--;

        // Check for throttling
        if (response.status === 429) {
          const retryAfter = response.headers.get('Retry-After');
          const waitTime = retryAfter ? parseInt(retryAfter) * 1000 : delay;
          
          console.warn(`Rate limited. Waiting ${waitTime}ms before retry.`);
          await new Promise(resolve => setTimeout(resolve, waitTime));
          
          retryCount++;
          delay = Math.min(delay * authConfig.retry.backoffMultiplier, authConfig.retry.maxDelay);
          continue;
        }

        // Check for server errors that should be retried
        if (response.status >= 500 && response.status < 600) {
          console.warn(`Server error ${response.status}. Retrying...`);
          await new Promise(resolve => setTimeout(resolve, delay));
          
          retryCount++;
          delay = Math.min(delay * authConfig.retry.backoffMultiplier, authConfig.retry.maxDelay);
          continue;
        }

        return response;
      } catch (error) {
        this.activeRequests--;
        
        if (retryCount < maxRetries - 1) {
          console.warn(`Request failed: ${error.message}. Retrying...`);
          await new Promise(resolve => setTimeout(resolve, delay));
          
          retryCount++;
          delay = Math.min(delay * authConfig.retry.backoffMultiplier, authConfig.retry.maxDelay);
        } else {
          throw error;
        }
      }
    }

    this.activeRequests--;
    throw new Error(`Request failed after ${maxRetries} attempts`);
  }

  generateCorrelationId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  async makeRequest(path, options = {}) {
    await this.initialize();

    try {
      let request = this.client.api(path);

      // Apply common query parameters
      if (options.select) {
        request = request.select(options.select);
      }
      if (options.top) {
        request = request.top(options.top);
      }
      if (options.filter) {
        request = request.filter(options.filter);
      }
      if (options.orderby) {
        request = request.orderby(options.orderby);
      }
      if (options.expand) {
        request = request.expand(options.expand);
      }
      if (options.search) {
        request = request.search(options.search);
      }

      // Execute the appropriate method
      switch (options.method?.toUpperCase()) {
        case 'POST':
          return await request.post(options.body);
        case 'PATCH':
          return await request.patch(options.body);
        case 'PUT':
          return await request.put(options.body);
        case 'DELETE':
          return await request.delete();
        default:
          return await request.get();
      }
    } catch (error) {
      this.handleGraphError(error);
    }
  }

  async makeBatchRequest(requests) {
    if (requests.length > 20) {
      throw new Error('Batch requests are limited to 20 operations');
    }

    await this.initialize();

    const batchContent = {
      requests: requests.map((req, index) => ({
        id: String(index + 1),
        method: req.method || 'GET',
        url: req.url,
        body: req.body,
        headers: req.headers,
      })),
    };

    try {
      const response = await this.client.api('/$batch').post(batchContent);
      return response.responses;
    } catch (error) {
      this.handleGraphError(error);
    }
  }

  handleGraphError(error) {
    const errorDetails = {
      message: error.message,
      code: error.code,
      statusCode: error.statusCode,
      correlationId: error.headers?.['client-request-id'] || 'unknown',
      timestamp: new Date().toISOString(),
    };

    if (error.body?.error) {
      errorDetails.innerError = error.body.error;
    }

    console.error('Graph API Error:', JSON.stringify(errorDetails, null, 2));

    // Enhanced error message for common scenarios
    if (error.statusCode === 401) {
      throw new Error('Authentication failed. Please re-authenticate.');
    } else if (error.statusCode === 403) {
      throw new Error('Insufficient permissions. Please check your app permissions.');
    } else if (error.statusCode === 404) {
      throw new Error('Resource not found. Please check the request path.');
    } else if (error.code === 'InvalidAuthenticationToken') {
      throw new Error('Invalid or expired token. Please re-authenticate.');
    }

    throw error;
  }

  // Utility methods for common operations
  async getWithSelect(path, fields) {
    return this.makeRequest(path, { select: fields.join(',') });
  }

  async postWithRetry(path, body) {
    return this.makeRequest(path, { method: 'POST', body });
  }

  async patchWithRetry(path, body) {
    return this.makeRequest(path, { method: 'PATCH', body });
  }

  async deleteWithRetry(path) {
    return this.makeRequest(path, { method: 'DELETE' });
  }

  // Pagination helper
  async *iterateAllPages(path, options = {}) {
    let nextLink = null;
    
    do {
      const response = nextLink 
        ? await this.client.api(nextLink).get()
        : await this.makeRequest(path, options);
      
      yield response.value || [];
      nextLink = response['@odata.nextLink'];
    } while (nextLink);
  }
}