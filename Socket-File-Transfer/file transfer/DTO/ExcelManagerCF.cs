namespace file_transfer
{
    using file_transfer.DTO;
    using System;
    using System.Data.Entity;
    using System.Linq;

    public class ExcelManagerCF : DbContext
    {
        // Your context has been configured to use a 'ExcelManagerCF' connection string from your application's 
        // configuration file (App.config or Web.config). By default, this connection string targets the 
        // 'file_transfer.ExcelManagerCF' database on your LocalDb instance. 
        // 
        // If you wish to target a different database and/or database provider, modify the 'ExcelManagerCF' 
        // connection string in the application configuration file.
        public ExcelManagerCF()
            : base("name=ExcelManagerCF")
        {
            Database.SetInitializer<ExcelManagerCF>(new CreateDB());
        }
        public virtual DbSet<Teacher> teachers { get; set; }
        public virtual DbSet<Room> rooms { get; set; }

        // Add a DbSet for each entity type that you want to include in your model. For more information 
        // on configuring and using a Code First model, see http://go.microsoft.com/fwlink/?LinkId=390109.

        // public virtual DbSet<MyEntity> MyEntities { get; set; }
    }

    //public class MyEntity
    //{
    //    public int Id { get; set; }
    //    public string Name { get; set; }
    //}
    public class CreateDB : CreateDatabaseIfNotExists<ExcelManagerCF>
    {
        protected override void Seed(ExcelManagerCF context)
        {
            context.teachers.Add(new Teacher
            {
                id = -1,
                name = "YYY",
                birthday = "1998/06/27",
                university = "DHBK",
            });
            context.rooms.Add(new Room
            {
                id = -1,
                nameRoom = "ZZZ",
                localtion = "DHNN",
                note = "Not note"
            });
            context.SaveChanges();
        }
    }
}