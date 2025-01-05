import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './user-session-revoke.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.USER_SESSION_REVOKE, () => {
  const userId = 'abcd1234-de71-4623-b4af-96380a352509';
  const userName = 'john.doe@contoso.com';
  const userNameWithDollar = "$john.doe@contoso.com";

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let promptIssued: boolean;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_SESSION_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      userName: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither userId nor userName is provided', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('prompts before revoking all sign-in sessions when confirm option not passed', async () => {
    await command.action(logger, { options: { userId: userId } });

    assert(promptIssued);
  });

  it('aborts revoking all sign-in sessions when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { userId: userId } });
    assert(postStub.notCalled);
  });

  it('revokes all sign-in sessions for a user specified by userId without prompting for confirmation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/revokeSignInSessions`) {
        return {
          value: true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, force: true, verbose: true } });
    assert(loggerLogSpy.calledOnceWith({ value: true }));
  });

  it('revokes all sign-in sessions for a user specified by UPN while prompting for confirmation', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/revokeSignInSessions`) {
        return {
          value: true
        };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { userName: userName } });
    assert(postRequestStub.calledOnce);
  });

  it('revokes all sign-in sessions for a user specified by UPN which starts with $ without prompting for confirmation', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userNameWithDollar}')/revokeSignInSessions`) {
        return {
          value: true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userNameWithDollar, force: true, verbose: true } });
    assert(postRequestStub.calledOnce);
  });

  it('handles error when user specified by userId was not found', async () => {
    sinon.stub(request, 'post').rejects({
      error:
      {
        code: 'Request_ResourceNotFound',
        message: `Resource '${userId}' does not exist or one of its queried reference-property objects are not present.`
      }
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(
      command.action(logger, { options: { userId: userId } }),
      new CommandError(`Resource '${userId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});