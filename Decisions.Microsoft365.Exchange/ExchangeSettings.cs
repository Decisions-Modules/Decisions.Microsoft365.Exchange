using System.ComponentModel;
using System.Runtime.Serialization;
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

namespace Decisions.Microsoft365.Exchange
{
    [ORMEntity("exchange_settings")]
    [Writable]
    public class ExchangeSettings : AbstractModuleSettings, INotifyPropertyChanged, IValidationSource
    {
        public ExchangeSettings()
        {
            this.EntityName = "Exchange Settings";
        }
        
        [ORMField]
        private string graphUrl = "https://graph.microsoft.com/v1.0";

        [PropertyClassification(0, "Graph URL", "Exchange Settings")]
        [DataMember]
        [WritableValue]
        public string GraphUrl
        {
            get => graphUrl;
            set
            {
                graphUrl = value.TrimEnd('/', '\\');
                OnPropertyChanged(nameof(GraphUrl));
            }
        }

        [ORMField]
        private string tokenId;

        [WritableValue]
        [PropertyClassification(new string[] { "Credentials" }, "OAuth Token", 1)]
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
                actions.Add(new EditEntityAction(typeof(ExchangeSettings), "Edit", "Edits Exchange Module Settings")
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