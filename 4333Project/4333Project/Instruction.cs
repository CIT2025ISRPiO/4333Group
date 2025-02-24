using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _4333Project
{
    public class Instruction
    {
        public Delegate Callable;
        public object[] args { get; set; }
        public void Execute() {
            this.Callable.DynamicInvoke(args);
        }
    }
}
