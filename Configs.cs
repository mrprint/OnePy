using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Xml;

namespace OnePy
{

#if (ONEPY45)
    public partial class OnePy45
#else
#if (ONEPY35)
    public partial class OnePy35
#endif
#endif
    {

        struct RuntimeConfig
        {
            static public List<string> additional_paths = new List<string>();
            static public List<string> additional_asms = new List<string>();
        }

        void LoadConfiguration()
        {
            Configuration cfg = ConfigurationManager.OpenExeConfiguration(Path.Combine(AssemblyDirectory, c_AddinName +".dll"));
            var libpaths = ((OnePyConfigSection)cfg.GetSection("onepy")).libpaths;
            foreach (ValueConfigElement ce in libpaths)
            {
                RuntimeConfig.additional_paths.Add(Path.GetFullPath(Environment.ExpandEnvironmentVariables(ce.Value)));
            }
            var load = ((OnePyConfigSection)cfg.GetSection("onepy")).load;
            foreach (ValueConfigElement ce in load)
            {
                RuntimeConfig.additional_asms.Add(ce.Value);
            }
        }

    }

    public class ValueConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("value", IsRequired = true)]
        public string Value
        {
            get
            {
                return this["value"] as string;
            }
        }
    }

    public class ValueConfigElementCollection : ConfigurationElementCollection
    {
        public ValueConfigElement this[int index]
        {
            get
            {
                return base.BaseGet(index) as ValueConfigElement;
            }
            set
            {
                if (base.BaseGet(index) != null)
                {
                    base.BaseRemoveAt(index);
                }
                this.BaseAdd(index, value);
            }
        }
        protected override ConfigurationElement CreateNewElement()
        {
            return new ValueConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ValueConfigElement)element).Value;
        }

        [ConfigurationProperty("name", IsKey = true)]
        public string Name
        {
            get
            {
                return this["name"] as string;
            }
        }

        [ConfigurationProperty("disabled", IsKey = true)]
        public bool Disabled
        {
            get
            {
                return (bool)this["disabled"];
            }
        }
    }

    public class OnePyConfigSection : ConfigurationSection
    {
        [ConfigurationProperty("libpaths")]
        public ValueConfigElementCollection libpaths
        {
            get 
            {
                return ((ValueConfigElementCollection)(base["libpaths"])); 
            }
        }

        [ConfigurationProperty("load")]
        public ValueConfigElementCollection load
        {
            get
            {
                return ((ValueConfigElementCollection)(base["load"]));
            }
        }
    }
}
