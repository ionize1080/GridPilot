using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
namespace SheetPilot.Core
{
    // Trie-based matching avoids comparing every cell with every identity in large workbooks.
    public sealed class IdentityMatcher
    {
        private sealed class Node { public Dictionary<char,Node> Next=new Dictionary<char,Node>();public string Replacement; }
        private readonly Node root=new Node();
        public IdentityMatcher(IDictionary<string,string> map)
        {foreach(var pair in map){var node=root;foreach(char c in pair.Key){char key=char.ToUpperInvariant(c);if(!node.Next.TryGetValue(key,out var next))node.Next[key]=next=new Node();node=next;}node.Replacement=pair.Value;}}
        public bool Contains(string text,CancellationToken token) => Walk(text,false,token)!=null;
        public string Replace(string text,CancellationToken token) => Walk(text,true,token)??text;
        private string Walk(string text,bool replace,CancellationToken token)
        {
            if(string.IsNullOrEmpty(text))return null;
            StringBuilder output=replace?new StringBuilder():null;bool found=false;
            for(int i=0;i<text.Length;i++)
            {
                if((i&255)==0)token.ThrowIfCancellationRequested();var node=root;string value=null;int end=i;
                for(int j=i;j<text.Length && node.Next.TryGetValue(char.ToUpperInvariant(text[j]),out node);j++)
                { if((j&255)==0)token.ThrowIfCancellationRequested();if(node.Replacement!=null){value=node.Replacement;end=j;} }
                if(value==null){output?.Append(text[i]);continue;}
                if(!replace)return "match";found=true;output.Append(value);i=end;
            }
            return found?output?.ToString():null;
        }
    }
}
