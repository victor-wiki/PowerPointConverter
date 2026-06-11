namespace PowerPointConverter.Builder
{
    public class StyleBuilder
    {
        private Dictionary<string, string> dict = new Dictionary<string, string>();

        public int Count => this.dict.Count;

        public void Add(string content)
        {
            if(!string.IsNullOrEmpty(content))
            {
                var items = content.Split(';');

                foreach(var item in items)
                {
                    var subItems = item.Split(':');

                    if(subItems.Length == 2)
                    {
                        this.Add(subItems[0].Trim(), subItems[1].Trim());
                    }
                }
            }
        }

        public bool Contains(string key)
        {
            return this.dict.ContainsKey(key);
        }

        public void Add(string key, string value)
        {
            if(this.dict.ContainsKey(key))
            {
                this.dict[key] = value;
            }
            else
            {
                this.dict.Add(key, value);
            }
        }

        public void Append(string key, string value)
        {
            if (this.dict.ContainsKey(key))
            {
                this.dict[key] += " " + value;
            }
            else
            {
                this.dict.Add(key, value);
            }
        }

        public void Remove(string key)
        {
            if(this.dict.ContainsKey(key))
            {
                this.dict.Remove(key);
            }
        }

        public override string ToString()
        {
            return string.Join(";", this.dict.Select(item => $"{item.Key}:{item.Value}"));
        }
    }
}
