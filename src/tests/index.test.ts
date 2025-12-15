import logger from '../utils/logger';

jest.mock('../utils/logger');

describe('src/index.ts', () => {
  it('should construct providers and create module providers', async () => {
    const { default: DgDataProviderAzureDevOps } = await import('../index');
    const provider = new DgDataProviderAzureDevOps('https://dev.azure.com/org/', 'token', '5.1', 'jfrog');

    expect(logger.info).toHaveBeenCalled();

    await expect(provider.getMangementDataProvider()).resolves.toBeDefined();
    await expect(provider.getTicketsDataProvider()).resolves.toBeDefined();
    await expect(provider.getGitDataProvider()).resolves.toBeDefined();
    await expect(provider.getPipelinesDataProvider()).resolves.toBeDefined();
    await expect(provider.getTestDataProvider()).resolves.toBeDefined();
    await expect(provider.getResultDataProvider()).resolves.toBeDefined();
    await expect(provider.getJfrogDataProvider()).resolves.toBeDefined();
  });

  it('should default jfrogToken to empty string when undefined', async () => {
    const { default: DgDataProviderAzureDevOps } = await import('../index');
    const provider = new DgDataProviderAzureDevOps('https://dev.azure.com/org/', 'token', '5.1');

    const jfrog = await provider.getJfrogDataProvider();
    expect(jfrog).toBeDefined();
  });
});
