import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './homesite-set.js';

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const siteUrl = `https:\\contoso.sharepoint.com\sites\Work`;
  const outputDefaultResponse = `The Home site has been set to ${siteUrl}. It may take some time for the change to apply. Check aka.ms/homesites for details.`;
  const defaultResponse = {
    "value": outputDefaultResponse
  };

  const outputVivaConnectionDefaultResponse = `The Home site has been set to ${siteUrl} and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details.`;
  const vivaConnectionDefaultResponse = {
    "value": outputVivaConnectionDefaultResponse
  };

  const outputErrorResponse = `[Error ID: 09149788-0a26-4cee-a333-699b81f629d7] The provided site url can't be set as a Home site. Check aka.ms\homesites for cmdlet requirements.`;
  const errorResponse = {
    error: {
      "odata.error": {
        "code": "-2147213238, Microsoft.SharePoint.SPException",
        "message": {
          "lang": "en-US",
          "value": outputErrorResponse
        }
      }
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets the specified site as the Home Site', async () => {
    const requestBody = { sphSiteUrl: siteUrl };
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/SetSPHSite`) {
        return defaultResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        verbose: true
      }
    } as any);
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('sets the specified site as the Home Site and sets the Viva Connections default experience to True', async () => {
    const requestBody = { sphSiteUrl: siteUrl, configuration: { vivaConnectionsDefaultStart: true } };
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/SetSPHSiteWithConfiguration`) {
        return vivaConnectionDefaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        vivaConnectionsDefaultStart: true
      }
    } as any);
    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles error when setting the Home Site', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw errorResponse;
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: siteUrl
      }
    } as any), new CommandError(outputErrorResponse));
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
