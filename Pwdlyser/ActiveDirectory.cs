using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Windows;

namespace ActiveDirectory
{
    /// <summary>
    /// Active Directory User.
    /// </summary>
    public class ADUser
    {
        #region constants

        /// <summary>
        /// Property name of sAM account name.
        /// </summary>
        public const string SamAccountNameProperty = "sAMAccountName";
        public const string DescriptionProperty = "description";
        public const string LastLoginProperty = "lastLogon";
        public const string PwdLastSetProperty = "pwdLastSet";
        public const string UserAccountControlProperty = "userAccountControl";


        /// <summary>
        /// Property name of canonical name.
        /// </summary>
        public const string CanonicalNameProperty = "CN";

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the canonical name of the user.
        /// </summary>
        public string CN { get; set; }

        /// <summary>
        /// Gets or sets the sAM account name
        /// </summary>
        public string SamAcountName { get; set; }

        public long LastLogin { get; set; }
        public long PwdLastSet { get; set; }
        public string Description { get; set; }
        public string UserAccountControl { get; set; }
        #endregion

        /// <summary>
        /// Gets all users of a given domain.
        /// </summary>
        /// <param name="domain">Domain to query. Should be given in the form ldap://domain.com/ </param>
        /// <returns>A list of users.</returns>
        public static List<ADUser> GetUsers(string domain, string username, string password)
        {
            List<ADUser> users = new List<ADUser>();

            if (username.Any())
            {
                using (DirectoryEntry searchRoot = new DirectoryEntry(domain, username, password))
                using (DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot))
                {
                    // Set the filter
                    directorySearcher.Filter = "(&(objectCategory=person)(objectClass=user))";

                    // Set the properties to load.
                    directorySearcher.PropertiesToLoad.Add(CanonicalNameProperty);
                    directorySearcher.PropertiesToLoad.Add(SamAccountNameProperty);
                    directorySearcher.PropertiesToLoad.Add(LastLoginProperty);
                    directorySearcher.PropertiesToLoad.Add(PwdLastSetProperty);
                    directorySearcher.PropertiesToLoad.Add(DescriptionProperty);
                    directorySearcher.PropertiesToLoad.Add(UserAccountControlProperty);

                    try
                    {
                        using (SearchResultCollection searchResultCollection = directorySearcher.FindAll())
                        {
                            foreach (SearchResult searchResult in searchResultCollection)
                            {
                                // Create new ADUser instance
                                var user = new ADUser();

                                // Set CN if available.
                                if (searchResult.Properties[CanonicalNameProperty].Count > 0)
                                    user.CN = searchResult.Properties[CanonicalNameProperty][0].ToString();

                                // Set sAMAccountName if available
                                if (searchResult.Properties[SamAccountNameProperty].Count > 0)
                                    user.SamAcountName = searchResult.Properties[SamAccountNameProperty][0].ToString();

                                // Description
                                if (searchResult.Properties[DescriptionProperty].Count > 0)
                                {
                                    user.Description = searchResult.Properties[DescriptionProperty][0].ToString();
                                    if (user.Description.Length >= 65)
                                    {
                                        user.Description = user.Description.Substring(0, 60) + "...";
                                    }
                                }

                                // Last Login
                                if (searchResult.Properties[LastLoginProperty].Count > 0)
                                    user.LastLogin = (long)searchResult.Properties[LastLoginProperty][0];

                                // Password Last Set
                                if (searchResult.Properties[PwdLastSetProperty].Count > 0)
                                    user.PwdLastSet = (long)searchResult.Properties[PwdLastSetProperty][0];

                                // Password Last Set
                                if (searchResult.Properties[UserAccountControlProperty].Count > 0)
                                {
                                    /*
                                        sEnabled = 'Enabled'
                                        s512 = 'Enabled Account'
                                        s514 = 'Disabled Account'
                                        s544 = 'Enabled, Password Not Required'
                                        s546 = 'Disabled, Password Not Required'
                                        s66048 = 'Account Enabled, Password Doesn\'t Expire'
                                        s66050 = 'Disabled, Password Doesn\'t Expire'
                                        s66080 = 'Enabled, Password Doesn\'t Expire & Not Required'
                                        s66082 = 'Disabled, Password Doesn\'t Expire & Not Required'
                                        s262656 = 'Enabled, Smartcard Required'
                                        s262658 = 'Disabled, Smartcard Required'
                                        s262688	= 'Enabled, Smartcard Required, Password Not Required'
                                        s262690 = 'Disabled, Smartcard Required, Password Not Required'
                                        s328192 = 'Enabled, Smartcard Required, Password Doesn\'t Expire'
                                        s328194 = 'Disabled, Smartcard Required, Password Doesn\'t Expire'
                                        s328224 = 'Enabled, Smartcard Required, Password Doesn\'t Expire & Not Required'
                                        s328226 = 'Disabled, Smartcard Required, Password Doesn\'t Expire & Not Required'
                                        sDisabled = 'Disabled'
                                     * */
                                    string UacTemp = searchResult.Properties[UserAccountControlProperty][0].ToString();
                                    switch (UacTemp)
                                    {
                                        case "Enabled":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "Disabled":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "512":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "514":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "544":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "546":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "66048":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "66050":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "66080":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "66082":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "262656":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "262658":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "262688":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "262690":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "328192":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "328194":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "328224":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "328226":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "131072":
                                            user.UserAccountControl = "Enabled";
                                            break;

                                        case "262144":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        default:
                                            user.UserAccountControl = "Enabled";
                                            break;
                                    }



                                }
                                // Add user to users list.
                                users.Add(user);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "Error: Cannot Retrieve User Information");
                    }
                }
            }
            else
            {


                using (DirectoryEntry searchRoot = new DirectoryEntry(domain))
                using (DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot))
                {
                    // Set the filter
                    directorySearcher.Filter = "(&(objectCategory=person)(objectClass=user))";

                    // Set the properties to load.
                    directorySearcher.PropertiesToLoad.Add(CanonicalNameProperty);
                    directorySearcher.PropertiesToLoad.Add(SamAccountNameProperty);
                    directorySearcher.PropertiesToLoad.Add(LastLoginProperty);
                    directorySearcher.PropertiesToLoad.Add(PwdLastSetProperty);
                    directorySearcher.PropertiesToLoad.Add(DescriptionProperty);
                    directorySearcher.PropertiesToLoad.Add(UserAccountControlProperty);

                    try
                    {
                        using (SearchResultCollection searchResultCollection = directorySearcher.FindAll())
                        {
                            foreach (SearchResult searchResult in searchResultCollection)
                            {
                                // Create new ADUser instance
                                var user = new ADUser();

                                // Set CN if available.
                                if (searchResult.Properties[CanonicalNameProperty].Count > 0)
                                    user.CN = searchResult.Properties[CanonicalNameProperty][0].ToString();

                                // Set sAMAccountName if available
                                if (searchResult.Properties[SamAccountNameProperty].Count > 0)
                                    user.SamAcountName = searchResult.Properties[SamAccountNameProperty][0].ToString();

                                // Description
                                if (searchResult.Properties[DescriptionProperty].Count > 0)
                                    user.Description = searchResult.Properties[DescriptionProperty][0].ToString();

                                // Last Login
                                if (searchResult.Properties[LastLoginProperty].Count > 0)
                                    user.LastLogin = (long)searchResult.Properties[LastLoginProperty][0];

                                // Password Last Set
                                if (searchResult.Properties[PwdLastSetProperty].Count > 0)
                                    user.PwdLastSet = (long)searchResult.Properties[PwdLastSetProperty][0];

                                // Password Last Set
                                if (searchResult.Properties[UserAccountControlProperty].Count > 0)
                                {
                                    /*
                                        sEnabled = 'Enabled'
                                        s512 = 'Enabled Account'
                                        s514 = 'Disabled Account'
                                        s544 = 'Enabled, Password Not Required'
                                        s546 = 'Disabled, Password Not Required'
                                        s66048 = 'Account Enabled, Password Doesn\'t Expire'
                                        s66050 = 'Disabled, Password Doesn\'t Expire'
                                        s66080 = 'Enabled, Password Doesn\'t Expire & Not Required'
                                        s66082 = 'Disabled, Password Doesn\'t Expire & Not Required'
                                        s262656 = 'Enabled, Smartcard Required'
                                        s262658 = 'Disabled, Smartcard Required'
                                        s262688	= 'Enabled, Smartcard Required, Password Not Required'
                                        s262690 = 'Disabled, Smartcard Required, Password Not Required'
                                        s328192 = 'Enabled, Smartcard Required, Password Doesn\'t Expire'
                                        s328194 = 'Disabled, Smartcard Required, Password Doesn\'t Expire'
                                        s328224 = 'Enabled, Smartcard Required, Password Doesn\'t Expire & Not Required'
                                        s328226 = 'Disabled, Smartcard Required, Password Doesn\'t Expire & Not Required'
                                        sDisabled = 'Disabled'
                                     * */
                                    string UacTemp = searchResult.Properties[UserAccountControlProperty][0].ToString();
                                    switch (UacTemp)
                                    {
                                        case "Enabled":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "Disabled":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "512":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "514":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "544":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "546":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "66048":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "66050":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "66080":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "66082":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "262656":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "262658":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "262688":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "262690":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "328192":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "328194":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        case "328224":
                                            user.UserAccountControl = "Enabled";
                                            break;
                                        case "328226":
                                            user.UserAccountControl = "Disabled";
                                            break;
                                        default:
                                            user.UserAccountControl = "Enabled";
                                            break;
                                    }



                                }
                                // Add user to users list.
                                users.Add(user);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "Error: Cannot Retrieve User Information");
                    }
                }
            }

            // Return all found users.
            return users;
        }
    }
}