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
import command from './commandset-add.js';

describe(commands.COMMANDSET_ADD, () => {
  let commandInfo: CommandInfo;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const validTitle = 'CLI Custom Action';
  const validClientSideComponentId = 'b206e130-1a5b-4ae7-86a7-4f91c9924d0a';
  const validWebUrl = 'https://contoso.sharepoint.com';
  const validClientSideComponentProperties = '{"testMessage":"Test message"}';
  const validListType = 'List';
  const commandactionResponse = {
    ClientSideComponentId: "b206e130-1a5b-4ae7-86a7-4f91c9924d0a",
    ClientSideComponentProperties: "",
    CommandUIExtension: null,
    Description: null,
    Group: null,
    HostProperties: "",
    Id: "680ccc51-7ddf-4dda-8696-fc606480cc3f",
    ImageUrl: null,
    Location: "ClientSideExtension.ListViewCommandSet.CommandBar",
    Name: null,
    RegistrationId: "100",
    RegistrationType: 0,
    Rights: {
      High: "0",
      Low: "0"
    },
    Scope: 2,
    ScriptBlock: null,
    ScriptSrc: null,
    Sequence: 0,
    Title: "CLI Custom Action",
    Url: null,
    VersionOfUserCustomAction: "16.0.1.0"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COMMANDSET_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        title: validTitle,
        listType: validListType,
        clientSideComponentId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listType is not valid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        title: validTitle,
        listType: 'Invalid list type',
        clientSideComponentId: validClientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is not valid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        title: validTitle,
        listType: validListType,
        clientSideComponentId: validClientSideComponentId,
        scope: 'Invalid scope'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if location is not valid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        title: validTitle,
        listType: validListType,
        clientSideComponentId: validClientSideComponentId,
        location: 'Invalid location'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if web url is not valid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'Invalid web url',
        title: validTitle,
        listType: validListType,
        clientSideComponentId: validClientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options specified', async () => {
    const actual = await command.validate({ options: { title: validTitle, webUrl: validWebUrl, listType: validListType, clientSideComponentId: validClientSideComponentId, scope: 'Web', location: 'Both', clientSideComponentProperties: validClientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds commandset with scope Web, list type list and location Both', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions`)) {
        {
          return commandactionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: validWebUrl, title: validTitle, clientSideComponentId: validClientSideComponentId, clientSideComponentProperties: validClientSideComponentProperties, listType: validListType, scope: 'Web', location: 'Both' } });
    assert(loggerLogSpy.calledWith(commandactionResponse));
  });

  it('adds commandset with scope Web and list type library', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions`)) {
        {
          return commandactionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, title: validTitle, clientSideComponentId: validClientSideComponentId, clientSideComponentProperties: validClientSideComponentProperties, listType: 'Library' } });
    assert(loggerLogSpy.calledWith(commandactionResponse));
  });

  it('adds commandset with location ContextMenu and listType SitePages', async () => {
    const response = commandactionResponse;
    response.Location = 'ClientSideExtension.ListViewCommandSet.ContextMenu';

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions`)) {
        {
          return response;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, title: validTitle, clientSideComponentId: validClientSideComponentId, clientSideComponentProperties: validClientSideComponentProperties, scope: 'Site', listType: 'SitePages', location: 'ContextMenu' } });
    assert(loggerLogSpy.calledWith(response));
  });

  it('adds commandset with location CommandBar', async () => {
    const response = commandactionResponse;
    response.Location = 'ClientSideExtension.ListViewCommandSet.CommandBar';

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions`)) {
        {
          return response;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, title: validTitle, clientSideComponentId: validClientSideComponentId, clientSideComponentProperties: validClientSideComponentProperties, scope: 'Site', listType: validListType, location: 'CommandBar' } });
    assert(loggerLogSpy.calledWith(response));
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong adding the commandset`
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions`) {
        throw error;
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: validWebUrl, title: validTitle, clientSideComponentId: validClientSideComponentId, clientSideComponentProperties: validClientSideComponentProperties, scope: 'Site' } } as any),
      new CommandError(`Something went wrong adding the commandset`));
  });

  it('offers autocomplete for the listType option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--listType') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the scope option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--scope') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the location option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--location') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });
});
