namespace PowerPointConverter.Builder
{
    public class StyleBuilder
    {
        private Dictionary<string, string> dict = new Dictionary<string, string>();

        public Dictionary<string, string> KeyValues => this.dict;

        public int Count => this.dict.Count;

        public StyleBuilder Add(string content)
        {
            var keyValues = GetKeyValues(content);

            foreach(var kp in keyValues)
            {
                if(!this.dict.ContainsKey(kp.Key))
                {
                    this.dict.Add(kp.Key, kp.Value);                       
                }
                else
                {
                    this.dict[kp.Key] = kp.Value;
                }
            }

            return this;
        }

        public static Dictionary<string, string> GetKeyValues(string content)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            if (!string.IsNullOrEmpty(content))
            {
                var items = content.Split(';');

                foreach (var item in items)
                {
                    var subItems = item.Split(':');

                    if (subItems.Length == 2)
                    {
                        dict.Add(subItems[0].Trim(), subItems[1].Trim());
                    }
                }
            }

            return dict;
        }

        public bool Contains(string key)
        {
            return this.dict.ContainsKey(key);
        }

        public StyleBuilder Add(string key, string value)
        {
            if (this.dict.ContainsKey(key))
            {
                this.dict[key] = value;
            }
            else
            {
                this.dict.Add(key, value);
            }

            return this;
        }

        public StyleBuilder Append(string key, string value)
        {
            if (this.dict.ContainsKey(key))
            {
                this.dict[key] += " " + value;
            }
            else
            {
                this.dict.Add(key, value);
            }

            return this;
        }

        public StyleBuilder Remove(string key)
        {
            if (this.dict.ContainsKey(key))
            {
                this.dict.Remove(key);
            }

            return this;
        }

        public string Get(string key)
        {
            if (key != null && this.dict.ContainsKey(key))
            {
                return this.dict[key];
            }

            return null;
        }

        public void Clear()
        {
            this.dict.Clear();
        }

        public override string ToString()
        {
            return string.Join(";", this.dict.Select(item => $"{item.Key}:{item.Value}"));
        }
    }
}
