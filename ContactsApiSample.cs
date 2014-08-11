/* Release under GPL V3 License 
 * The GPL (V2 or V3) is a copyleft license that requires anyone who modifies or updates this code 
 * to make the source available under the same terms.
 */

using Microsoft.Office365.Exchange;
using Microsoft.Office365.OAuth;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace O365APIContactsSample
{
    public class MyContact
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Email { get; set; }

        public string JobTitle { get; set; }

        public byte[] Picture { get; set; }
    }

    public static class ContactsAPISample
    {
        const string ServiceResourceId = "https://outlook.office365.com";
        static readonly Uri ServiceEndpointUri = new Uri("https://outlook.office365.com/ews/odata");

        static string _lastLoggedInUser;
        static DiscoveryContext _discoveryContext;
        static ExchangeClient _exchangeClient;

        public static async Task Create(MyContact myContact)
        {
            //create the Exchange contact
            IContact contact = new Contact();
            contact.GivenName = myContact.Name;
            contact.DisplayName = myContact.Name;
            contact.EmailAddress1 = myContact.Email;
            contact.JobTitle = myContact.JobTitle;

            //add the contact
            await _exchangeClient.Me.Contacts.AddContactAsync(contact);

            //initialize the contact picture entity
            IFileAttachment contactPhotoAttachment = new FileAttachment();
            contactPhotoAttachment.IsContactPhoto = true;
            contactPhotoAttachment.ContentBytes = myContact.Picture;
            contactPhotoAttachment.Name = "Picture";

            //add the picture attachment
            await ((IContactFetcher)contact).Attachments.AddAttachmentAsync(contactPhotoAttachment);

            //update the contact
            await contact.UpdateAsync();          
        }

        public static async Task<ObservableCollection<MyContact>> GetContacts()
        {
            ObservableCollection<MyContact> myContactList = new ObservableCollection<MyContact>();

            //query contacts
            var contactsResults = await _exchangeClient.Me.Contacts.OrderBy(c=>c.DisplayName).ExecuteAsync();

            //fetches the first 50 contacts
            //default page count is 50
            var contacts = contactsResults.CurrentPage.ToList();

            //enumerate and build your contact object
            foreach (var contact in contacts )
            {
                MyContact myContact = new MyContact();

                myContact.Id = contact.Id;
                myContact.Name = String.Format("{0} {1}", contact.GivenName, contact.Surname);
                myContact.Email = contact.EmailAddress1;
                myContact.Picture = await GetContactImage(contact.Id);

                myContactList.Add(myContact);
            }

            return myContactList;
        }      

        public static async Task<byte[]> GetContactImage(string strContactId)
        {
            byte[] bytesContactPhoto = new byte[0];

            //query the contact by contact Id
            //GetById is broken so you need to use this instead
            var contact = await _exchangeClient.Me.Contacts[strContactId].ExecuteAsync();

            if (contact != null)
            {
                //get the attachments
                var attachmentResult = await ((IContactFetcher)contact).Attachments.ExecuteAsync();
                var attachments = attachmentResult.CurrentPage.ToArray();

                //find the attachment entity that is a contact photo
                var contactPhotoAttachment = attachments.OfType<IFileAttachment>().FirstOrDefault(a => a.IsContactPhoto);

                if (contactPhotoAttachment != null)
                {
                   //read the photo bytes
                   bytesContactPhoto = contactPhotoAttachment.ContentBytes;
                }
            }

            return bytesContactPhoto;
        }

        public static async Task<bool> Delete(string strContactId)
        {
            bool bStatus = false;

            //query the contact by contact Id
            //GetById is broken so you need to use this instead
            var contact = await _exchangeClient.Me.Contacts[strContactId].ExecuteAsync();

            if (contact != null)
            {
                //delete the contact
                await contact.DeleteAsync();

                bStatus = true;
            }

            return bStatus;
        }

        public static async Task<bool> Update(MyContact myContact)
        {
            bool bStatus = false;

            //query the contact by contact Id
            //GetById is broken so you need to use this instead
            var contact = await _exchangeClient.Me.Contacts[myContact.Id].ExecuteAsync();

            if (contact != null)
            {
                contact.EmailAddress1 = myContact.Email;
                contact.JobTitle = myContact.JobTitle;

                //update the contact
                await contact.UpdateAsync();

                bStatus = true;
            }

            return bStatus;
        }

        public static async Task SignIn()
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            var dcr = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId);

            _lastLoggedInUser = dcr.UserId;

            _exchangeClient = new ExchangeClient(ServiceEndpointUri, async () =>
            {
                return (await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId, _discoveryContext.AppIdentity.ClientId, new Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier(dcr.UserId, Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifierType.UniqueId))).AccessToken;
            });
        }

        public static async Task SignOut()
        {
            if (string.IsNullOrEmpty(_lastLoggedInUser))
            {
                return;
            }

            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            await _discoveryContext.LogoutAsync(_lastLoggedInUser);
        }
    }
}
