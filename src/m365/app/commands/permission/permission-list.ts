import { Application, AppRole, AppRoleAssignment, OAuth2PermissionGrant, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import AppCommand, { appCommandOptions } from '../../../base/AppCommand.js';
import commands from '../../commands.js';

interface ApiPermission {
  resource: string;
  permission: string;
  type: string;
}

interface ServicePrincipalInfo {
  appId?: string;
  id?: string;
}

enum GetServicePrincipal {
  withPermissions,
  withPermissionDefinitions
}

class AppPermissionListCommand extends AppCommand {
  public get name(): string {
    return commands.PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists API permissions for the current Microsoft Entra app';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return appCommandOptions;
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const servicePrincipal = await this.getServicePrincipal({ appId: this.appId }, logger, GetServicePrincipal.withPermissions);

      let permissions: ApiPermission[];
      if (servicePrincipal) {
        // service principal found, get permissions from the service principal
        permissions = await this.getServicePrincipalPermissions(servicePrincipal, logger);
      }
      else {
        // service principal not found, get permissions from app registration
        permissions = await this.getAppRegPermissions(this.appId as string, logger);
      }

      await logger.log(permissions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getServicePrincipal(servicePrincipalInfo: ServicePrincipalInfo, logger: Logger, mode: GetServicePrincipal): Promise<ServicePrincipal | undefined> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving service principal ${servicePrincipalInfo.appId ?? servicePrincipalInfo.id}`);
    }

    const lookupUrl: string = servicePrincipalInfo.appId ? `?$filter=appId eq '${servicePrincipalInfo.appId}'&` : `/${servicePrincipalInfo.id}?`;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals${lookupUrl}$select=appId,id,displayName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: ServicePrincipal[] } | ServicePrincipal>(requestOptions);

    if ((servicePrincipalInfo.id && !response) ||
      (servicePrincipalInfo.appId && (response as { value: ServicePrincipal[] }).value.length === 0)) {
      return undefined;
    }

    const servicePrincipal = servicePrincipalInfo.appId ?
      (response as { value: ServicePrincipal[] }).value[0] :
      response as ServicePrincipal;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving permissions for service principal ${servicePrincipal.id}...`);
    }

    const permissionsPromises = [];

    switch (mode) {
      case GetServicePrincipal.withPermissions:
        const appRoleAssignmentsRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoleAssignments`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        const oauth2PermissionGrantsRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oauth2PermissionGrants`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        permissionsPromises.push(...[
          request.get<{ value: AppRoleAssignment[] }>(appRoleAssignmentsRequestOptions),
          request.get<{ value: OAuth2PermissionGrant[] }>(oauth2PermissionGrantsRequestOptions)
        ]);
        break;
      case GetServicePrincipal.withPermissionDefinitions:
        const oauth2PermissionScopesRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oauth2PermissionScopes`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        const appRolesRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoles`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        permissionsPromises.push(...[
          request.get<{ value: PermissionScope[] }>(oauth2PermissionScopesRequestOptions),
          request.get<{ value: AppRole[] }>(appRolesRequestOptions)
        ]);
        break;
    }

    const permissions = await Promise.all(permissionsPromises);

    switch (mode) {
      case GetServicePrincipal.withPermissions:
        servicePrincipal.appRoleAssignments = permissions[0].value;
        servicePrincipal.oauth2PermissionGrants = permissions[1].value as any;
        break;
      case GetServicePrincipal.withPermissionDefinitions:
        servicePrincipal.oauth2PermissionScopes = permissions[0].value as any;
        servicePrincipal.appRoles = permissions[1].value as any;
        break;
    }

    return servicePrincipal;
  }

  private async getServicePrincipalPermissions(servicePrincipal: ServicePrincipal, logger: Logger): Promise<ApiPermission[]> {
    if (this.verbose) {
      await logger.logToStderr(`Resolving permissions for the service principal...`);
    }

    const apiPermissions: ApiPermission[] = [];

    // hash table for resolving resource IDs to names
    const resourceLookup: { [key: string]: string } = {};
    // list of service principals for which to load permissions
    const servicePrincipalsToResolve: ServicePrincipalInfo[] = [];

    const appRoleAssignments = servicePrincipal.appRoleAssignments as AppRoleAssignment[];
    apiPermissions.push(...appRoleAssignments.map(appRoleAssignment => {
      // store resource name for resolving OAuth2 grants
      resourceLookup[appRoleAssignment.resourceId as string] = appRoleAssignment.resourceDisplayName as string;
      // add to the list of service principals to load to get the app role
      // display name
      if (!servicePrincipalsToResolve.find(r => r.id === appRoleAssignment.resourceId)) {
        servicePrincipalsToResolve.push({ id: appRoleAssignment.resourceId as string });
      }

      return {
        resource: appRoleAssignment.resourceDisplayName as string,
        // we store the app role ID temporarily and will later resolve to display name
        permission: appRoleAssignment.appRoleId as string,
        type: 'Application'
      };
    }));

    const oauth2Grants = servicePrincipal.oauth2PermissionGrants as OAuth2PermissionGrant[];

    oauth2Grants.forEach(oauth2Grant => {
      // see if we can resolve the resource name from the resources
      // retrieved from app role assignments
      const resource = resourceLookup[oauth2Grant.resourceId as string] ?? oauth2Grant.resourceId as string;
      if (resource === oauth2Grant.resourceId as string &&
        !servicePrincipalsToResolve.find(r => r.id === oauth2Grant.resourceId)) {
        // resource name not found in the resources
        // add it to the list of resources to resolve
        servicePrincipalsToResolve.push({ id: oauth2Grant.resourceId as string });
      }

      const scopes = (oauth2Grant.scope as string).split(' ');
      scopes.forEach(scope => {
        apiPermissions.push({
          resource,
          permission: scope,
          type: 'Delegated'
        });
      });
    });

    if (servicePrincipalsToResolve.length > 0) {
      const servicePrincipals = await Promise
        .all(servicePrincipalsToResolve
          .map(servicePrincipalInfo => this.getServicePrincipal(servicePrincipalInfo, logger, GetServicePrincipal.withPermissionDefinitions) as ServicePrincipal));
      servicePrincipals.forEach(servicePrincipal => {
        apiPermissions.forEach(apiPermission => {
          if (apiPermission.resource === servicePrincipal.id) {
            apiPermission.resource = servicePrincipal.displayName as string;
          }

          if (apiPermission.resource === servicePrincipal.displayName &&
            apiPermission.type === 'Application') {
            apiPermission.permission = (servicePrincipal.appRoles as AppRole[])
              .find(appRole => appRole.id === apiPermission.permission)?.value as string ?? apiPermission.permission;
          }
        });
      });
    }

    return apiPermissions;
  }

  private async getAppRegistration(appId: string, logger: Logger): Promise<Application> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the Entra application registration with appId '${appId}'`);
    }

    const app: Application = await entraApp.getAppRegistrationByAppId(appId);

    return app;
  }

  private async getAppRegPermissions(appId: string, logger: Logger): Promise<ApiPermission[]> {
    const application = await this.getAppRegistration(appId, logger);

    if ((application.requiredResourceAccess as RequiredResourceAccess[]).length === 0) {
      return [];
    }

    const servicePrincipalsToResolve: ServicePrincipalInfo[] =
      (application.requiredResourceAccess as RequiredResourceAccess[])
        .map(resourceAccess => {
          return {
            appId: resourceAccess.resourceAppId as string
          };
        });
    const servicePrincipals = await Promise
      .all(servicePrincipalsToResolve.map(servicePrincipalInfo =>
        this.getServicePrincipal(servicePrincipalInfo, logger, GetServicePrincipal.withPermissionDefinitions) as ServicePrincipal));

    const apiPermissions: ApiPermission[] = [];
    (application.requiredResourceAccess as RequiredResourceAccess[]).forEach(requiredResourceAccess => {
      const servicePrincipal = servicePrincipals
        .find(servicePrincipal => servicePrincipal?.appId === requiredResourceAccess.resourceAppId as string);
      const resourceName = servicePrincipal?.displayName as string ?? requiredResourceAccess.resourceAppId as string;
      (requiredResourceAccess.resourceAccess as ResourceAccess[]).forEach(permission => {
        apiPermissions.push({
          resource: resourceName,
          permission: this.getPermissionName(permission.id as string, permission.type as string, servicePrincipal),
          type: permission.type === 'Role' ? 'Application' : 'Delegated'
        });
      });
    });

    return apiPermissions;
  }

  private getPermissionName(permissionId: string, permissionType: string, servicePrincipal: ServicePrincipal | undefined): string {
    if (!servicePrincipal) {
      return permissionId;
    }

    switch (permissionType) {
      case 'Role':
        return (servicePrincipal.appRoles as AppRole[])
          .find(appRole => appRole.id === permissionId)?.value as string ?? permissionId;
      case 'Scope':
        return (servicePrincipal.oauth2PermissionScopes as PermissionScope[])
          .find(permissionScope => permissionScope.id === permissionId)?.value as string ?? permissionId;
    }
    /* c8 ignore next 4 */
    // permissionType is either 'Scope' or 'Role' but we need a safe default
    // to avoid building errors. This code will never be reached.
    return permissionId;
  }
}

export default new AppPermissionListCommand();