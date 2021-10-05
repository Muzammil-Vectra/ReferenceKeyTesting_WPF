using System;
using Inventor;
using Environment = System.Environment;
using File = System.IO.File;

namespace ReferenceKeyTesting_WPF
{
    public class ReferenceKeyManagerClass
    {
        public ReferenceKeyManager ReferenceKeyManager;
        public int KeyContextPointer;
        public ReferenceKeyManagerClass(Document document)
        {
            ReferenceKeyManager = document.ReferenceKeyManager;
        }

        public void CreateKeyContextOnce()
        {
            KeyContextPointer = ReferenceKeyManager.CreateKeyContext();
        } 
        public string GetReferenceKeyAndKeyContext(dynamic entity, out string keyContext)
        {
            keyContext = null;
            try
            {
                if (ReferenceKeyManager != null && entity != null)
                {
                    byte[] referenceKey = new byte[] { };
                    byte[] keyContextArray = new byte[] { };
                    entity.GetReferenceKey(ref referenceKey, KeyContextPointer);
                    string key = ReferenceKeyManager.KeyToString(ref referenceKey);
                    ReferenceKeyManager.SaveContextToArray(KeyContextPointer, ref keyContextArray);
                    keyContext = ReferenceKeyManager.KeyToString(ref keyContextArray);
                    return key;
                }
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
            return null;
        }
        public string GetReferenceKeyOnly(dynamic entity, out string keyContext)
        {
            keyContext = "";
            if (entity != null)
            {
                byte[] referenceKey = new byte[] { };
                entity.GetReferenceKey(ref referenceKey, KeyContextPointer);
                string key = ReferenceKeyManager.KeyToString(ref referenceKey);
                return key;
            }
            return null;
        }

        public string GetKeyContextOnly()
        {
            byte[] keyContextArray = new byte[] { };
            ReferenceKeyManager.SaveContextToArray(KeyContextPointer, ref keyContextArray);
            string keyContext = ReferenceKeyManager.KeyToString(ref keyContextArray);
            return keyContext;
        }
        public string GetKeyContext()
        {
            byte[] keyContextArray = new byte[] { };
            ReferenceKeyManager.SaveContextToArray(KeyContextPointer, ref keyContextArray);
            string keyContext = ReferenceKeyManager.KeyToString(ref keyContextArray);
            //ReferenceKeyManager.ReleaseKeyContext(KeyContextPointer);
            return keyContext;
        }
        public int LoadKeyContext(string keyContextString)
        {
            byte[] keyContextArray = new byte[] { };
            ReferenceKeyManager.StringToKey(keyContextString, ref keyContextArray);
            int  keyContext = ReferenceKeyManager.LoadContextFromArray(keyContextArray);
            return keyContext;
        }
        public dynamic GetEntityFromReferenceKey(string referenceKey, int keyContext)
        {
            if (ReferenceKeyManager != null)
            {
                byte[] key = new byte[] { };
                object entity=null;
                ReferenceKeyManager.StringToKey(referenceKey, ref key);
                entity = ReferenceKeyManager.BindKeyToObject(ref key, keyContext, out object matchType);
                return entity;
            }
            return null;
        }
    }
}
