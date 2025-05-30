import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listItemId: number;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
  force?: boolean;
}

class SpoListItemRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from list item permissions';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--listItemId <listItemId>'
      },
      {
        option: '--principalId [principalId]'
      },
      {
        option: '--upn [upn]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        if (args.options.listItemId && isNaN(args.options.listItemId)) {
          return `Specified listItemId ${args.options.listItemId} is not a number`;
        }

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoleAssignment(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from listitem ${args.options.listItemId} from list ${args.options.listId || args.options.listTitle} from site ${args.options.webUrl}?` });

      if (result) {
        await this.removeRoleAssignment(logger, args.options);
      }
    }
  }

  private async removeRoleAssignment(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing role assignment from listitem in site at ${options.webUrl}...`);
    }

    try {
      let requestUrl: string = `${options.webUrl}/_api/web/`;
      if (options.listId) {
        requestUrl += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')/`;
      }
      else if (options.listTitle) {
        requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
      }

      requestUrl += `items(${options.listItemId})/`;

      if (options.upn) {
        options.principalId = await this.getUserPrincipalId(options, logger);
      }
      else if (options.groupName) {
        options.principalId = await this.getGroupPrincipalId(options, logger);
      }
      else if (options.entraGroupId || options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = options.entraGroupId
          ? await entraGroup.getGroupById(options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
        options.principalId = siteUser.Id;
      }

      await this.removeRoleAssignmentWithRequestUrl(requestUrl, logger, options);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeRoleAssignmentWithRequestUrl(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: any = {
      url: `${requestUrl}roleassignments/removeroleassignment(principalid='${options.principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }
}

export default new SpoListItemRoleAssignmentRemoveCommand();