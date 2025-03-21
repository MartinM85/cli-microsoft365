import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand.js';
import commands from '../../commands.js';
import { ToDoTask } from '../../ToDoTask.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listName?: string;
  listId?: string;
}

class TodoTaskGetCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Get a specific task from a Microsoft To Do task list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listName: typeof args.options.listName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--listName [listName]'
      },
      {
        option: '--listId [listId]'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listName'] });
  }

  private async getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return args.options.listId;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.listName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ id: string }] }>(requestOptions);

    const taskList = response.value[0];
    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listId: string = await this.getTodoListId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/todo/lists/${listId}/tasks/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const item: ToDoTask = await request.get(requestOptions);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(item);
      }
      else {
        await logger.log({
          id: item.id,
          title: item.title,
          status: item.status,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TodoTaskGetCommand();