using System;
using System.Collections.Generic; // Тут вообще не используется этот нэймспэйс
using System.Linq; // И этот
using System.Text;
using System.Threading.Tasks // И этот тоже

namespace WorkerToExel
{
    public class Worker
    {
        public byte[] email { get; set; } // Все свойства должны придерживаться нотации PascalCase, CamelCase подходит только для параметров и переменных
        public byte[] lname { get; set; } // Сокращение наименований - плохая практика
        public byte[] fname { get; set; }
        public byte[] password { get; set; }
    }
}
