using System;
using System.Collections.Generic;
using System.Text;

namespace OpenOfficeUploader.Entities
{
    public class Person
    {
        public static Person Create(string surName, string firstName, string phone, string email, DateTime dateOfBirth, string middleName = default)
            => new Person(surName, firstName, phone, email, dateOfBirth, middleName);
        private Person(string surName, string firstName, string phone, string email, DateTime dateOfBirth, string middleName = default)
        {
            if(string.IsNullOrWhiteSpace(surName)) throw new ArgumentNullException(nameof(surName));
            if(string.IsNullOrWhiteSpace(firstName)) throw new ArgumentNullException(nameof(firstName));
            if(string.IsNullOrWhiteSpace(phone)) throw new ArgumentNullException(nameof(phone));
            if(string.IsNullOrWhiteSpace(email)) throw new ArgumentNullException(nameof(email));
            SurName = surName;
            FirstName = firstName;
            MiddleName = middleName;
            Phone = phone;
            Email = email;
            DateOfBirth = dateOfBirth;
            Id = Guid.NewGuid();
        }
        private Person() { }
        public Guid Id { get; set; }
        public string SurName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }
    }
}
