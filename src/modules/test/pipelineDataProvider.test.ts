import DgDataProviderAzureDevOps from '../..';

require('dotenv').config();
jest.setTimeout(60000);

const orgUrl = process.env.ORG_URL;
const token = process.env.PAT;
const dgDataProviderAzureDevOps = new DgDataProviderAzureDevOps(orgUrl, token);

describe('pipeline module - tests', () => {
  test('should return pipeline info', async () => {
    let pipelinesDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await pipelinesDataProvider.getPipelineBuildByBuildId('tests', 244);
    expect(json.id).toBe(244);
  });
  test('should return Release definition', async () => {
    let pipelinesDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await pipelinesDataProvider.GetReleaseByReleaseId('tests', 1);
    expect(json.id).toBe(1);
  });
  test('should return OK(200) as response ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let result = await PipelineDataProvider.TriggerBuildById(
      'tests',
      '14',
      '{"test":"param1","age":"26","name":"denis" }'
    );
    expect(result.status).toBe(200);
  });
  test('should the path to zip file as response ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let result = await PipelineDataProvider.GetArtifactByBuildId(
      'tests',
      '245', //buildId
      '_tests' //artifactName
    );
    expect(result).toBeDefined();
  });

  test('should return pipeline run history ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await PipelineDataProvider.GetPipelineRunHistory('tests', '14');
    expect(json).toBeDefined();
  });

  test('should return release defenition history ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await PipelineDataProvider.GetReleaseHistory('tests', '1');
    expect(json).toBeDefined();
  });

  test('should return all pipelines ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await PipelineDataProvider.GetAllPipelines('tests');
    expect(json).toBeDefined();
  });

  test('should return all releaseDefenitions ', async () => {
    let PipelineDataProvider = await dgDataProviderAzureDevOps.getPipelinesDataProvider();
    let json = await PipelineDataProvider.GetAllReleaseDefenitions('tests');
    expect(json).toBeDefined();
  });
});
