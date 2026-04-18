import { Calendar } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';

const calendarColors = ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  name: z.string().optional().alias('n'),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  calendarGroupId: z.string().optional(),
  calendarGroupName: z.string().optional(),
  color: z.enum(calendarColors).optional(),
  isDefault: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDAR_SET;
  }

  public get description(): string {
    return 'Updates a calendar for a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both.'
      })
      .refine(options => !(options.calendarGroupId && options.calendarGroupName), {
        error: 'Specify either calendarGroupId or calendarGroupName, but not both.'
      })
      .refine(options => [options.name, options.color, options.isDefault].filter(o => o !== undefined).length > 0, {
        error: 'Specify at least one of the following options: name, color, or isDefault.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      let requestUrl: string;

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }

        const userIdentifier = args.options.userId ?? args.options.userName;
        requestUrl = this.buildRequestUrl(userIdentifier!);
      }
      else {
        if (args.options.userId || args.options.userName) {
          const currentUserId = accessToken.getUserIdFromAccessToken(token);
          const currentUserName = accessToken.getUserNameFromAccessToken(token);
          const isOtherUser = (args.options.userId && args.options.userId !== currentUserId) ||
            (args.options.userName && args.options.userName.toLowerCase() !== currentUserName?.toLowerCase());

          if (isOtherUser) {
            const scopes = accessToken.getScopesFromAccessToken(token);
            const hasSharedScope = scopes.some(s => s === 'Calendars.ReadWrite.Shared');

            if (!hasSharedScope) {
              throw 'To update calendars of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.';
            }
          }

          const userIdentifier = args.options.userId ?? args.options.userName;
          requestUrl = this.buildRequestUrl(userIdentifier!);
        }
        else {
          requestUrl = this.buildRequestUrl(undefined);
        }
      }

      if (this.verbose) {
        await logger.logToStderr(`Updating calendar '${args.options.id}'...`);
      }

      let calendarGroupId = args.options.calendarGroupId;

      if (args.options.calendarGroupName) {
        const userIdForGroup = args.options.userId ?? args.options.userName ?? accessToken.getUserIdFromAccessToken(token);
        const calendarGroupResult = await calendarGroup.getUserCalendarGroupByName(userIdForGroup, args.options.calendarGroupName, 'id');
        calendarGroupId = calendarGroupResult.id!;
      }

      const url = this.buildCalendarUrl(requestUrl, args.options.id, calendarGroupId);

      const requestOptions: CliRequestOptions = {
        url,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: this.createRequestBody(args)
      };

      const result = await request.patch<Calendar>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private buildRequestUrl(userIdentifier: string | undefined): string {
    if (userIdentifier) {
      return `${this.resource}/v1.0/users('${userIdentifier}')`;
    }

    return `${this.resource}/v1.0/me`;
  }

  private buildCalendarUrl(baseUrl: string, calendarId: string, calendarGroupId: string | undefined): string {
    if (calendarGroupId) {
      return `${baseUrl}/calendarGroups/${calendarGroupId}/calendars/${calendarId}`;
    }

    return `${baseUrl}/calendars/${calendarId}`;
  }

  private createRequestBody(args: CommandArgs): any {
    const data: any = {};

    if (args.options.name !== undefined) {
      data.name = args.options.name;
    }

    if (args.options.color !== undefined) {
      data.color = args.options.color;
    }

    if (args.options.isDefault !== undefined) {
      data.isDefaultCalendar = args.options.isDefault;
    }

    return data;
  }
}

export default new OutlookCalendarSetCommand();
