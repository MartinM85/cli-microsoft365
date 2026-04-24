import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';

export const options = z.strictObject({ ...globalOptionsZod.shape });

class ExternalConnectionListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists external connections defined in the Microsoft Search';
  }

  public alias(): string[] | undefined {
    return [commands.EXTERNALCONNECTION_LIST];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'state'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const connections = await odata.getAllItems(`${this.resource}/v1.0/external/connections`);
      await logger.log(connections);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionListCommand();