using System.ComponentModel;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
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
    [Writable]
    public class Exchange365Settings : AbstractModuleSettings, INotifyPropertyChanged, IValidationSource
    {
        public Exchange365Settings()
        {
            this.EntityName = "Exchange 365 Settings";
        }

        [ORMField]
        private string tokenId;

        [WritableValue]
        [PropertyClassification(new string[] { "Credentials" }, "OAuth Token", 0)]
        [TokenPicker]
        public string TokenId
        {
            get => tokenId;
            set
            {
                tokenId = value;
                OnPropertyChanged(nameof(TokenId));
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