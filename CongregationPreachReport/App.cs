using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CongregationPreachReport
{
    public class App
    {
        private readonly IConfiguration config;

        public App(IConfiguration config)
        {
            this.config = config;
        }

        public void Run()
        {
            Console.WriteLine("Hello World ");

            Console.WriteLine("ReleaseMode: " + config.GetValue<bool>("ReleaseMode"));

            Console.WriteLine("Directory: " + Directory.GetCurrentDirectory());

            Console.ReadLine();
        }
    }
}
