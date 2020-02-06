using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace G1ANT.Addon.Mscrm
{
    public class MsCrmManager
    {
        private readonly static MsCrmManager instance = new MsCrmManager();
        private MsCrmManager() { }
        public static MsCrmManager Instance => instance;
        private static List<MsCrmWrapper> launchedCRM = new List<MsCrmWrapper>();

        public MsCrmWrapper CurrentCRM { get; private set; }

        public bool SwitchCRM(int id)
        {
            var tmpCRM = launchedCRM.FirstOrDefault(x => x.Id == id);
            CurrentCRM = tmpCRM ?? CurrentCRM;
            return tmpCRM != null;
        }

        public List<MsCrmWrapper> GetLaunchedCRM()
        {
            return launchedCRM;
        }

        private int GetNextId()
        {
            return launchedCRM.Any<MsCrmWrapper>() ? launchedCRM.Max(x => x.Id) + 1 : 0;
        }

        public MsCrmWrapper AddCRM()
        {
            int assignedId = GetNextId();
            MsCrmWrapper wrapper = new MsCrmWrapper(assignedId);
            launchedCRM.Add(wrapper);
            CurrentCRM = wrapper;
            return wrapper;
        }

        public MsCrmWrapper AttachToExistingCRM(string name, string by, bool msCrmRecorder = false)
        {
            int assignedId = GetNextId();
            
            MsCrmWrapper wrapper = new MsCrmWrapper(assignedId, name, by, msCrmRecorder);
            launchedCRM.Add(wrapper);
            CurrentCRM = wrapper;
            return wrapper;
        }

        public void FindAnyActiveCRM()
        {
            int assignedId = GetNextId();

            IntPtr iHandle = FindWindow("✱Internet Explorer✱", 3000);
            if (iHandle != IntPtr.Zero)
            { 
                MsCrmWrapper wrapper = new MsCrmWrapper(assignedId, "crm", "url");
                if (wrapper != null)
                {
                    launchedCRM.Add(wrapper);
                    CurrentCRM = wrapper;
                }                
                else
                {
                    throw new ApplicationException("Specified CRM window not found");
                }
            }
            else
            {
                throw new ApplicationException("Specified CRM window not found");
            }
        }

        private IntPtr FindWindow(string name, int timeout)
        {
            Process[] list = Process.GetProcessesByName(name);
            if (list.Length > 0)
            {
                return list[0].Handle;
            }
            return new IntPtr(-1);
        }

        public void Detach(MsCrmWrapper wrapper)
        {
            var toRemove = launchedCRM.Where(x => x == wrapper).FirstOrDefault();
            if (toRemove != null)
            {
                launchedCRM.Remove(toRemove);
            }
            if (CurrentCRM == wrapper)
            {
                CurrentCRM = null;
            }
        }
    }

}
