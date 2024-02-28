using System.ComponentModel;
using System.Runtime.Serialization;
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
using Microsoft.Graph;

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
        
        [ORMField(typeof(StringArrayFieldConverter))]
        private string[] scopes;

        // Multi-tenant apps can use "common",
        // single-tenant apps must use the tenant ID from the Azure portal
        [ORMField]
        private string tenantId = "common";

        // Value from app registration
        [ORMField]
        private string clientId;
        
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
        
        [PropertyClassification(0, "Tenant ID", "Exchange 365 Settings")]
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
        
        [PropertyClassification(1, "Client ID", "Exchange 365 Settings")]
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