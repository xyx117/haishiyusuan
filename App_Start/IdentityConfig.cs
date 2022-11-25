using yusuanxiangmu.Models;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace yusuanxiangmu
{
    // This is useful if you do not want to tear down the database each time you run the application.
    // You want to create a new database if the Model changes
    // public class MyDbInitializer : DropCreateDatabaseIfModelChanges<MyDbContext> 这个适用于在开发阶段，经过改变模型时使用
    //CreateDatabaseIfNotExists方法会在没有数据库时创建一个，这是默认行为。
    //DropCreateDatabaseAlways  如果你想在每次运行时都重新生成数据库就可以用这个方法。

    public class MyDbInitializer : CreateDatabaseIfNotExists<MyUserDbContext>
    {
        protected override void Seed(MyUserDbContext context)
        {

            //AddUserAndRoles();
            InitializeIdentityForEF(context);
            base.Seed(context);
        }

        //Identity初始化 代码段
        private void InitializeIdentityForEF(MyUserDbContext context)
        {

            var UserManager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(context));
            var RoleManager = new RoleManager<IdentityRole>(new RoleStore<IdentityRole>(context));
          
            string name = "Admin";
            string password = "123456";

            string name2 = "处长";
            string name3 = "科长";
            string name4 = "工作人员";
            
            //Create Role Admin if it does not exist
            if (!RoleManager.RoleExists(name))
            {
                var roleresult = RoleManager.Create(new IdentityRole(name));

            }

            if (!RoleManager.RoleExists(name2))
            {
                var roleresult = RoleManager.Create(new IdentityRole(name2));

            }

            if (!RoleManager.RoleExists(name3))
            {
                var roleresult = RoleManager.Create(new IdentityRole(name3));

            }

            if (!RoleManager.RoleExists(name4))
            {
                var roleresult = RoleManager.Create(new IdentityRole(name4));

            }

              //var userexist = UserManager.FindByName("pass");
          
                //Create User=Admin with password=123456
            var user = new ApplicationUser() { UserName="admin",FirstName="lu",LastName="xian",Email="lu@hainnu.edu.cn"};
                    
            var adminresult = UserManager.Create(user, password);

            //Add User Admin to Role Admin
            if (adminresult.Succeeded)
            {
                var result = UserManager.AddToRole(user.Id, name);
            }
        }

    }
}