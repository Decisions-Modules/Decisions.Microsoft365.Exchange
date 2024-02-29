using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Runtime.Serialization.DataContracts;
using Azure.Identity;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Actions;
using DecisionsFramework.ServiceLayer.Actions.Common;
using DecisionsFramework.ServiceLayer.Services.Accounts;
using DecisionsFramework.ServiceLayer.Services.Administration;
using DecisionsFramework.ServiceLayer.Services.Folder;
using DecisionsFramework.ServiceLayer.Utilities;

namespace Decisions.Exchange365
{
    [ORMEntity("exchange365_settings")]
    [DataContract]
    [Writable]
    public class Exchange365Settings : AbstractModuleSettings, INotifyPropertyChanged, IValidationSource
    {
        public Exchange365Settings()
        {
            this.EntityName = "Exchange 365 Settings";
        }
        
        // Value from app registration
        [ORMField]
        private string clientId;
        
        // single-tenant apps must use the tenant ID from the Azure portal
        [ORMField]
        private string tenantId = "common";

        [ORMField(typeof(StringArrayFieldConverter))]
        private string[] scopes;
        
        [ORMField]
        private string clientSecret;

        [ORMField]
        private string authorizationCode;

        [PropertyClassification(0, "Client ID", "Exchange 365 Settings")]
        [DataMember]
        [WritableValue]
        public string ClientId
        {
            get => clientId;
            set
            {
                clientId = value;
                OnPropertyChanged(nameof(ClientId));
            }
        }
        
        [PropertyClassification(1, "Tenant ID", "Exchange 365 Settings")]
        [DataMember]
        [WritableValue]
        public string TenantId
        {
            get => tenantId;
            set
            {
                tenantId = value;
                OnPropertyChanged(nameof(TenantId));
            }
        }
        
        [PropertyClassification(2, "Scopes", "Exchange 365 Settings")]
        [DataMember]
        [WritableValue]
        public string[] Scopes
        {
            get => scopes;
            set
            {
                scopes = value;
                OnPropertyChanged(nameof(Scopes));
            }
        }
        
        [PropertyClassification(3, "Client Secret", "Exchange 365 Settings")]
        [DataMember]
        [WritableValue]
        public string ClientSecret
        {
            get => clientSecret;
            set
            {
                clientSecret = value;
                OnPropertyChanged(nameof(ClientSecret));
            }
        }
        
        [PropertyClassification(4, "Authorization Code", "Exchange 365 Settings")]
        [DataMember]
        [WritableValue]
        public string AuthorizationCode
        {
            get => authorizationCode;
            set
            {
                authorizationCode = value;
                OnPropertyChanged(nameof(AuthorizationCode));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ValidationIssue[] GetValidationIssues()
        {
            List<ValidationIssue> issues = new List<ValidationIssue>();

            return issues.ToArray();
        }

        public override BaseActionType[] GetActions(AbstractUserContext userContext, EntityActionType[] types)
        {
            List<BaseActionType> actions = new List<BaseActionType>();

            Account userAccount = userContext.GetAccount();

            FolderPermission permission = FolderService.GetAccountEffectivePermissionInternal(
                new SystemUserContext(), this.EntityFolderID, userAccount.AccountID);

            bool canAdministrate =
                FolderPermission.CanAdministrate == (FolderPermission.CanAdministrate & permission) ||
                userAccount.GetUserRights<PortalAdministratorModuleRight>() != null ||
                userAccount.IsAdministrator();

            if (canAdministrate)
            {
                actions.Add(new EditEntityAction(typeof(Exchange365Settings), "Edit", "Edits Exchange 365 Module Settings")
                {
                    IsDefaultGridAction = true,
                    OkActionName = "SAVE",
                    CancelActionName = null
                });
            }

            return actions.ToArray();
        }
    }
}