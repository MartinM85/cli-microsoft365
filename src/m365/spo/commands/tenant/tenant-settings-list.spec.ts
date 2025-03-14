import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-settings-list.js';

describe(commands.TENANT_SETTINGS_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles client.svc promise error', async () => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('handles error while getting tenant appcatalog', async () => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc"
          }
        ]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('lists the tenant settings (debug)', async () => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1218", "ErrorInfo": null, "TraceCorrelationId": "6148899e-a042-6000-ee90-5bfa05d08b79"
          }, 4, {
            "IsNull": false
          }, 5, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "6648899e-a042-6000-ee90-5bfa05d08b79|908bed80-a04a-4433-b4a0-883d9847d11d:ea1787c6-7ce2-4e71-be47-5e0deb30f9ee\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

            ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 1, "DefaultSharingLinkType": 1, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
              ""
            ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": true, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

            ], "RequireAcceptingAccountMatchInvitedAccount": true, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 66700, "ResourceQuotaAllocated": 13668, "RootSiteUrl": "https:\u002f\u002fprufinancial.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": "microsoft.com pramerica.ie pramericacdsdev.com prudential.com prufinancial.onmicrosoft.com", "SharingBlockedDomainList": "deloitte.com", "SharingCapability": 1, "SharingDomainRestrictionMode": 1, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": false, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 4448256, "StorageQuotaAllocated": 676508312, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": false, "HideDefaultThemes": true
          }
        ]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert.strictEqual(loggerLogSpy.lastCall.args[0].AllowDownloadingNonWebViewableFiles, true);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].BccExternalSharingInvitationsList, null);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].HideDefaultThemes, true);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].UserVoiceForFeedbackEnabled, false);
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["_ObjectType_"], undefined);
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["_ObjectIdentity_"], undefined);

    assert.strictEqual(loggerLogSpy.lastCall.args[0]["SharingCapability"], 'ExternalUserSharingOnly');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["SharingDomainRestrictionMode"], 'AllowList');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["ODBMembersCanShare"], 'Unspecified');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["ODBAccessRequests"], 'Unspecified');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["DefaultSharingLinkType"], 'Direct');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["FileAnonymousLinkType"], 'Edit');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["FolderAnonymousLinkType"], 'Edit');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["DefaultLinkPermission"], 'View');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["ConditionalAccessPolicy"], 'AllowFullAccess');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["SpecialCharactersStateInFileFolderNames"], 'Allowed');
    assert.strictEqual(loggerLogSpy.lastCall.args[0]["LimitedAccessFileType"], 'WebPreviewableFiles');
  });

  it('handles tenant settings error', async () => {
    // get tenant app catalog
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
          }
        ]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Timed out'));
  });
});
