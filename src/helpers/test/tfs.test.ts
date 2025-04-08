// First, mock axios BEFORE importing TFSServices
import axios from 'axios';
jest.mock('axios');

// Set up the mock axios instance that axios.create will return
const mockAxiosInstance = {
    request: jest.fn()
};
(axios.create as jest.Mock).mockReturnValue(mockAxiosInstance);

// NOW import TFSServices (it will use our mock)
import { TFSServices } from '../tfs';
import logger from '../../utils/logger';

// Mock logger
jest.mock('../../utils/logger');

describe('TFSServices', () => {
    // Store the original implementation of random to restore it later
    const originalRandom = Math.random;

    beforeEach(() => {
        jest.clearAllMocks();
        // Mock Math.random to return a predictable value for tests with retry
        Math.random = jest.fn().mockReturnValue(0.5);
    });

    afterEach(() => {
        // Restore the original Math.random implementation
        Math.random = originalRandom;
    });

    describe('downloadZipFile', () => {
        it('should download a zip file successfully', async () => {
            // Arrange
            const url = 'https://example.com/file.zip';
            const pat = 'token123';
            const mockResponse = { data: Buffer.from('zip-file-content') };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.downloadZipFile(url, pat);

            // Assert
            expect(result).toEqual(mockResponse);
            expect(mockAxiosInstance.request).toHaveBeenCalledWith({
                url,
                headers: { 'Content-Type': 'application/zip' },
                auth: { username: '', password: pat }
            });
        });

        it('should log and throw error when download fails', async () => {
            // Arrange
            const url = 'https://example.com/file.zip';
            const pat = 'token123';
            const mockError = new Error('Network error');

            // Configure mock to throw error
            mockAxiosInstance.request.mockRejectedValueOnce(mockError);

            // Act & Assert
            await expect(TFSServices.downloadZipFile(url, pat)).rejects.toThrow();
            expect(logger.error).toHaveBeenCalledWith(`error download zip file , url : ${url}`);
        });
    });

    describe('fetchAzureDevOpsImageAsBase64', () => {
        it('should fetch and convert image to base64', async () => {
            // Arrange
            const url = 'https://example.com/image.png';
            const pat = 'token123';
            const mockResponse = {
                data: Buffer.from('image-data'),
                headers: { 'content-type': 'image/png' }
            };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.fetchAzureDevOpsImageAsBase64(url, pat);

            // Assert
            expect(result).toEqual('data:image/png;base64,aW1hZ2UtZGF0YQ==');
            expect(mockAxiosInstance.request).toHaveBeenCalledWith(expect.objectContaining({
                url,
                method: 'get',
                auth: { username: '', password: pat },
                responseType: 'arraybuffer'
            }));
        });

        it('should handle errors and retry for retryable errors', async () => {
            // Arrange
            const url = 'https://example.com/image.png';
            const pat = 'token123';

            // Create a rate limit error (retry-eligible)
            const rateLimitError = new Error('Rate limit exceeded');
            (rateLimitError as any).response = { status: 429 };

            // Configure mock to fail once then succeed
            mockAxiosInstance.request
                .mockRejectedValueOnce(rateLimitError)
                .mockResolvedValueOnce({
                    data: Buffer.from('image-data'),
                    headers: { 'content-type': 'image/png' }
                });

            // Mock setTimeout to execute immediately
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act
            const result = await TFSServices.fetchAzureDevOpsImageAsBase64(url, pat);

            // Assert
            expect(result).toEqual('data:image/png;base64,aW1hZ2UtZGF0YQ==');
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(2);
            expect(logger.warn).toHaveBeenCalled();
        });
    });

    describe('getItemContent', () => {
        it('should get item content successfully with GET request', async () => {
            // Arrange
            const url = 'https://example.com/api/item';
            const pat = 'token123';
            const mockResponse = { data: { id: 123, name: 'Test Item' } };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.getItemContent(url, pat);

            // Assert
            expect(result).toEqual(mockResponse.data);
            expect(mockAxiosInstance.request).toHaveBeenCalledWith(expect.objectContaining({
                url: url.replace(/ /g, '%20'),
                method: 'get',
                auth: { username: '', password: pat },
                timeout: 10000 // Verify the actual timeout value is 10000ms
            }));
        });

        it('should fail when request times out', async () => {
            // Arrange
            const url = 'https://example.com/api/slow-item';
            const pat = 'token123';

            // Create a timeout error
            const timeoutError = new Error('timeout of 1000ms exceeded');
            timeoutError.name = 'TimeoutError';
            (timeoutError as any).code = 'ECONNABORTED';

            // Configure mock to simulate timeout (will retry 3 times by default)
            mockAxiosInstance.request
                .mockRejectedValueOnce(timeoutError)
                .mockRejectedValueOnce(timeoutError)
                .mockRejectedValueOnce(timeoutError);

            // Mock setTimeout to execute immediately for faster tests
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act & Assert
            await expect(TFSServices.getItemContent(url, pat)).rejects.toThrow('timeout');
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(3); // Initial + 2 retries
            expect(logger.warn).toHaveBeenCalledTimes(2); // Two retry warnings
        });

        it('should fail when network connection fails', async () => {
            // Arrange
            const url = 'https://example.com/api/item';
            const pat = 'token123';

            // Create different network errors
            const connectionResetError = new Error('socket hang up');
            (connectionResetError as any).code = 'ECONNRESET';

            const connectionRefusedError = new Error('connect ECONNREFUSED');
            (connectionRefusedError as any).code = 'ECONNREFUSED';

            // Configure mock to simulate different network failures on each retry
            mockAxiosInstance.request
                .mockRejectedValueOnce(connectionResetError)
                .mockRejectedValueOnce(connectionRefusedError)
                .mockRejectedValueOnce(connectionResetError);

            // Mock setTimeout to execute immediately for faster tests
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act & Assert
            await expect(TFSServices.getItemContent(url, pat)).rejects.toThrow();
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(2); // Initial + 2 retries
            expect(logger.error).toHaveBeenCalled(); // Should log detailed error
        });

        it('should handle DNS resolution failures', async () => {
            // Arrange
            const url = 'https://nonexistent-domain.example.com/api/item';
            const pat = 'token123';

            // Create DNS resolution error
            const dnsError = new Error('getaddrinfo ENOTFOUND nonexistent-domain.example.com');
            (dnsError as any).code = 'ENOTFOUND';

            // Configure mock to simulate DNS failure
            mockAxiosInstance.request.mockRejectedValue(dnsError);

            // Act & Assert
            await expect(TFSServices.getItemContent(url, pat)).rejects.toThrow('ENOTFOUND');
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(2); // Should retry DNS failures too
        });

        it('should handle spaces in URL by replacing them with %20', async () => {
            // Arrange
            const url = 'https://example.com/api/item with spaces';
            const pat = 'token123';
            const mockResponse = { data: { id: 123, name: 'Test Item' } };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            await TFSServices.getItemContent(url, pat);

            // Assert
            expect(mockAxiosInstance.request).toHaveBeenCalledWith(
                expect.objectContaining({ url: 'https://example.com/api/item%20with%20spaces' })
            );
        });


    });

    describe('getJfrogRequest', () => {
        it('should make a successful GET request to JFrog', async () => {
            // Arrange
            const url = 'https://jfrog.example.com/api/artifacts';
            const headers = { Authorization: 'Bearer token123' };
            const mockResponse = { data: { artifacts: [{ name: 'artifact1' }] } };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.getJfrogRequest(url, headers);

            // Assert
            expect(result).toEqual(mockResponse.data);
            expect(mockAxiosInstance.request).toHaveBeenCalledWith({
                url,
                method: 'GET',
                headers
            });
        });

        it('should handle errors from JFrog API', async () => {
            // Arrange
            const url = 'https://jfrog.example.com/api/artifacts';
            const headers = { Authorization: 'Bearer token123' };
            const mockError = new Error('JFrog API error');

            // Configure mock to throw error
            mockAxiosInstance.request.mockRejectedValueOnce(mockError);

            // Act & Assert
            await expect(TFSServices.getJfrogRequest(url, headers)).rejects.toThrow();
            expect(logger.error).toHaveBeenCalled();
        });
    });

    describe('postRequest', () => {
        it('should make a successful POST request', async () => {
            // Arrange
            const url = 'https://example.com/api/resource';
            const pat = 'token123';
            const data = { name: 'New Resource' };
            const mockResponse = { data: { id: 123, name: 'New Resource' } };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.postRequest(url, pat, 'post', data);

            // Assert
            expect(result).toEqual(mockResponse);
            expect(mockAxiosInstance.request).toHaveBeenCalledWith({
                url,
                method: 'post',
                auth: { username: '', password: pat },
                data,
                headers: { headers: { 'Content-Type': 'application/json' } }
            });
        });

        it('should work with custom headers and methods', async () => {
            // Arrange
            const url = 'https://example.com/api/resource';
            const pat = 'token123';
            const data = { name: 'Update Resource' };
            const customHeaders = { 'Content-Type': 'application/xml' };
            const mockResponse = { data: { id: 123, name: 'Updated Resource' } };

            // Configure mock response
            mockAxiosInstance.request.mockResolvedValueOnce(mockResponse);

            // Act
            const result = await TFSServices.postRequest(url, pat, 'put', data, customHeaders);

            // Assert
            expect(result).toEqual(mockResponse);
            expect(mockAxiosInstance.request).toHaveBeenCalledWith({
                url,
                method: 'put',
                auth: { username: '', password: pat },
                data,
                headers: customHeaders
            });
        });

        it('should handle errors in POST requests', async () => {
            // Arrange
            const url = 'https://example.com/api/resource';
            const pat = 'token123';
            const data = { name: 'New Resource' };
            const mockError = new Error('Validation error');
            (mockError as any).response = { status: 400, data: { message: 'Invalid data' } };

            // Configure mock to throw error
            mockAxiosInstance.request.mockRejectedValueOnce(mockError);

            // Act & Assert
            await expect(TFSServices.postRequest(url, pat, 'post', data)).rejects.toThrow();
            expect(logger.error).toHaveBeenCalled();
        });
    });

    describe('executeWithRetry', () => {
        it('should retry on network timeouts', async () => {
            // Arrange
            const url = 'https://example.com/api/slow-resource';
            const pat = 'token123';

            // Create a timeout error
            const timeoutError = new Error('timeout of 10000ms exceeded');

            // Configure mock to fail with timeout then succeed
            mockAxiosInstance.request
                .mockRejectedValueOnce(timeoutError)
                .mockResolvedValueOnce({ data: { success: true } });

            // Mock setTimeout to execute immediately
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act
            const result = await TFSServices.getItemContent(url, pat);

            // Assert
            expect(result).toEqual({ success: true });
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(2);
            expect(logger.warn).toHaveBeenCalled();
        });

        it('should retry on server errors (5xx)', async () => {
            // Arrange
            const url = 'https://example.com/api/unstable-resource';
            const pat = 'token123';

            // Create a 503 error
            const serverError = new Error('Service unavailable');
            (serverError as any).response = { status: 503 };

            // Configure mock to fail with server error then succeed
            mockAxiosInstance.request
                .mockRejectedValueOnce(serverError)
                .mockResolvedValueOnce({ data: { success: true } });

            // Mock setTimeout to execute immediately
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act
            const result = await TFSServices.getItemContent(url, pat);

            // Assert
            expect(result).toEqual({ success: true });
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(2);
        });

        it('should give up after max retry attempts', async () => {
            // Arrange
            const url = 'https://example.com/api/failing-resource';
            const pat = 'token123';

            // Create a persistent server error
            const serverError = new Error('Service unavailable');
            (serverError as any).response = { status: 503 };

            // Configure mock to always fail with server error
            mockAxiosInstance.request.mockRejectedValue(serverError);

            // Mock setTimeout to execute immediately
            jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                fn();
                return {} as any;
            });

            // Act & Assert
            await expect(TFSServices.getItemContent(url, pat)).rejects.toThrow();

            // Should try original request + retries up to maxAttempts (3)
            expect(mockAxiosInstance.request).toHaveBeenCalledTimes(3);
        });
    });
    // Add these test cases to the existing test file

    describe('Stress testing', () => {
        describe('getItemContent - Sequential large data requests', () => {
            it('should handle multiple sequential requests with large datasets', async () => {
                // Arrange
                const url = 'https://example.com/api/large-dataset';
                const pat = 'token123';

                // Create 5 large responses of different sizes
                const responses = Array(5).fill(0).map((_, i) => {
                    // Create increasingly large responses (500KB, 1MB, 1.5MB, 2MB, 2.5MB)
                    const size = 25000 * (i + 1);
                    return {
                        data: {
                            items: Array(size).fill(0).map((_, j) => ({
                                id: j,
                                name: `Item ${j}`,
                                description: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
                                data: `value-${j}-${Math.random().toString(36).substring(2, 15)}`
                            }))
                        }
                    };
                });

                // Configure the mock to return the different responses in sequence
                for (const response of responses) {
                    mockAxiosInstance.request.mockResolvedValueOnce(response);
                }

                // Act - Make multiple sequential requests
                const results = [];
                for (let i = 0; i < responses.length; i++) {
                    const result = await TFSServices.getItemContent(`${url}/${i}`, pat);
                    results.push(result);
                }

                // Assert
                expect(results.length).toBe(responses.length);
                for (let i = 0; i < results.length; i++) {
                    expect(results[i].items.length).toBe(responses[i].data.items.length);
                }
                expect(mockAxiosInstance.request).toHaveBeenCalledTimes(responses.length);
            });

            it('should handle sequential requests with mixed success/failure patterns', async () => {
                // Arrange
                const url = 'https://example.com/api/sequential';
                const pat = 'token123';

                // Set up a sequence of responses/errors
                // 1. Success
                // 2. Rate limit error (429) - should retry and succeed
                // 3. Success
                // 4. Server error (503) - should retry and succeed
                // 5. Success

                const rateLimitError = new Error('Rate limit exceeded');
                (rateLimitError as any).response = { status: 429 };

                const serverError = new Error('Server error');
                (serverError as any).response = { status: 503 };

                // Response 1
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 1, success: true }
                });

                // Response 2 - fails with rate limit first, then succeeds
                mockAxiosInstance.request
                    .mockRejectedValueOnce(rateLimitError)
                    .mockResolvedValueOnce({
                        data: { id: 2, success: true }
                    });

                // Response 3
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 3, success: true }
                });

                // Response 4 - fails with server error first, then succeeds
                mockAxiosInstance.request
                    .mockRejectedValueOnce(serverError)
                    .mockResolvedValueOnce({
                        data: { id: 4, success: true }
                    });

                // Response 5
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 5, success: true }
                });

                // Mock setTimeout to execute immediately for retry delay
                jest.spyOn(global, 'setTimeout').mockImplementation((fn: any) => {
                    fn();
                    return {} as any;
                });

                // Act - Make sequential requests
                const results = [];
                for (let i = 1; i <= 5; i++) {
                    const result = await TFSServices.getItemContent(`${url}/${i}`, pat);
                    results.push(result);
                }

                // Assert
                expect(results.length).toBe(5);
                expect(results.map(r => r.id)).toEqual([1, 2, 3, 4, 5]);

                // Check total number of requests (5 successful responses + 2 retries)
                expect(mockAxiosInstance.request).toHaveBeenCalledTimes(7);
            });
        });

        describe('postRequest - Sequential large data submissions', () => {
            it('should handle multiple sequential POST requests with large payloads', async () => {
                // Arrange
                const url = 'https://example.com/api/resource';
                const pat = 'token123';

                // Create 5 increasingly large payloads
                const payloads = Array(5).fill(0).map((_, i) => {
                    // Create payloads of increasing size (100KB, 200KB, 300KB, 400KB, 500KB)
                    const size = 5000 * (i + 1);
                    return {
                        name: `Resource ${i + 1}`,
                        description: `Large resource ${i + 1}`,
                        items: Array(size).fill(0).map((_, j) => ({
                            id: j,
                            value: `item-${j}-${Math.random().toString(36).substring(2, 15)}`,
                            timestamp: new Date().toISOString()
                        }))
                    };
                });

                // Set up mock responses
                for (let i = 0; i < payloads.length; i++) {
                    mockAxiosInstance.request.mockResolvedValueOnce({
                        data: { id: i + 1, status: 'created', itemCount: payloads[i].items.length }
                    });
                }

                // Act - Make multiple sequential POST requests
                const results = [];
                for (let i = 0; i < payloads.length; i++) {
                    const result = await TFSServices.postRequest(url, pat, 'post', payloads[i]);
                    results.push(result);
                }

                // Assert
                expect(results.length).toBe(payloads.length);
                for (let i = 0; i < results.length; i++) {
                    expect(results[i].data.id).toBe(i + 1);
                    expect(results[i].data.itemCount).toBe(payloads[i].items.length);
                }
                expect(mockAxiosInstance.request).toHaveBeenCalledTimes(payloads.length);
            });

            it('should handle a mix of success and failure during sequential POST operations', async () => {
                // Arrange
                const url = 'https://example.com/api/resource';
                const pat = 'token123';

                // Create test data
                const payloads = Array(5).fill(0).map((_, i) => ({ name: `Resource ${i + 1}` }));

                // Validation error
                const validationError = new Error('Validation error');
                (validationError as any).response = {
                    status: 400,
                    data: { message: 'Invalid data', details: 'Field X is required' }
                };

                // Server error
                const serverError = new Error('Server error');
                (serverError as any).response = { status: 500 };

                // Configure mock responses/errors for each request
                // 1. Success
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 1, status: 'created' }
                });

                // 2. Validation error
                mockAxiosInstance.request.mockRejectedValueOnce(validationError);

                // 3. Success
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 3, status: 'created' }
                });

                // 4. Server error
                mockAxiosInstance.request.mockRejectedValueOnce(serverError);

                // 5. Success
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 5, status: 'created' }
                });

                // Act & Assert
                // Request 1 - should succeed
                const result1 = await TFSServices.postRequest(url, pat, 'post', payloads[0]);
                expect(result1.data.id).toBe(1);

                // Request 2 - should fail with validation error
                await expect(TFSServices.postRequest(url, pat, 'post', payloads[1]))
                    .rejects.toThrow('Validation error');

                // Request 3 - should succeed despite previous failure
                const result3 = await TFSServices.postRequest(url, pat, 'post', payloads[2]);
                expect(result3.data.id).toBe(3);

                // Request 4 - should fail with server error
                await expect(TFSServices.postRequest(url, pat, 'post', payloads[3]))
                    .rejects.toThrow('Server error');

                // Request 5 - should succeed despite previous failure
                const result5 = await TFSServices.postRequest(url, pat, 'post', payloads[4]);
                expect(result5.data.id).toBe(5);

                // Verify all requests were made
                expect(mockAxiosInstance.request).toHaveBeenCalledTimes(5);
            });

            it('should handle large payload with many nested objects', async () => {
                // Arrange
                const url = 'https://example.com/api/complex-resource';
                const pat = 'token123';

                // Create a deeply nested object structure
                const createNestedObject = (depth: number, breadth: number, current = 0): any => {
                    if (current >= depth) {
                        return { value: `leaf-${Math.random()}` };
                    }

                    const children: Record<string, any> = {};
                    for (let i = 0; i < breadth; i++) {
                        children[`child-${current}-${i}`] = createNestedObject(depth, breadth, current + 1);
                    }

                    return {
                        id: `node-${current}-${Math.random()}`,
                        level: current,
                        children
                    };
                };

                // Create a complex payload with depth 5 and breadth 5 (5^5 = 3,125 nodes)
                const complexPayload = {
                    name: 'Complex Resource',
                    type: 'hierarchical',
                    rootNode: createNestedObject(5, 5)
                };

                // Configure mock response
                mockAxiosInstance.request.mockResolvedValueOnce({
                    data: { id: 123, status: 'created', complexity: 'high' }
                });

                // Act
                const result = await TFSServices.postRequest(url, pat, 'post', complexPayload);

                // Assert
                expect(result.data.id).toBe(123);
                expect(result.data.status).toBe('created');

                // Verify the request was made with the complex payload
                expect(mockAxiosInstance.request).toHaveBeenCalledWith(expect.objectContaining({
                    url,
                    method: 'post',
                    data: complexPayload
                }));
            });
        });

        describe('High volume sequential operations', () => {
            it('should handle a high volume of sequential API calls', async () => {
                // Arrange
                const url = 'https://example.com/api/items';
                const pat = 'token123';
                const requestCount = 50; // Make 50 sequential requests

                // Configure mock responses for all requests
                for (let i = 0; i < requestCount; i++) {
                    mockAxiosInstance.request.mockResolvedValueOnce({
                        data: { id: i, success: true, timestamp: new Date().toISOString() }
                    });
                }

                // Act
                const results = [];
                for (let i = 0; i < requestCount; i++) {
                    // Alternate between GET and POST requests
                    if (i % 2 === 0) {
                        const result = await TFSServices.getItemContent(`${url}/${i}`, pat);
                        results.push(result);
                    } else {
                        const result = await TFSServices.postRequest(url, pat, 'post', { itemId: i });
                        results.push(result.data);
                    }
                }

                // Assert
                expect(results.length).toBe(requestCount);
                for (let i = 0; i < requestCount; i++) {
                    expect(results[i].id).toBe(i);
                    expect(results[i].success).toBe(true);
                }
                expect(mockAxiosInstance.request).toHaveBeenCalledTimes(requestCount);
            });
        });
    });


});