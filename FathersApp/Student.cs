using System.Collections.Generic;

namespace FathersApp
{
    public class Student
    {
        private string SchoolId;
        private string StudentId;
        public Dictionary<string,string> ListOfTasks;

        public string SchoolId1 { get => SchoolId; set => SchoolId = value; }
        public string StudentId1 { get => StudentId; set => StudentId = value; }

        public Student(string school, string student, Dictionary<string, string> listoftasks)
        {
            this.SchoolId1 = school;
            this.StudentId1 = student;
            this.ListOfTasks = new Dictionary<string, string>(listoftasks); //before initialization of Student class user is should initialize Task collection.
        }
    }
}
