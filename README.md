# RTF-Exporter


A C# library to generate .RTF text files from any string object data, stylized and ready for any text processor. No fancy dependencies or restrictive licenses.


---


## Simple usage


IDisposable style:


```C#
using RTFExporter;

public class Example {

    public void IDisposableExample() {

        using (RTFDocument doc = new RTFDocument("example.rtf")) {
            var p = doc.AppendParagraph();
            var t = p.AppendText("Boy toy named Troy used to live in Detroit\n");
            t.content += "Big big big money, he was gettin' some coins";

            t.style.bold = true;
            t.style.color = new Color(255, 0, 0);
            t.style.fontFamily = "Courier";
        }

    }

}
```


String style:


```C#
using RTFExporter;

public class Example {

    public void StringExample() {

        RTFDocument doc = new RTFDocument();
        RTFParagraph p = new RTFParagraph(doc);

        RTFText t1 = new RTFText(p, "My anaconda don't, my anaconda don't\n");
        t1.SetStyle(new Color(255, 0, 0), 18, "Helv");

        RTFText t2 = new RTFText(p, "My anaconda don't want none unless you got buns, hun");
        t2.style = t1.style;

        string output = RTFParser.ToString(doc);

    }

}
```


---


## Why?  


This repository originally was made to export text files from Unity game engine, but can be used for any purpose.


---


## License  (WTFPL-2.0)


<a href="http://www.wtfpl.net/"><img src="http://www.wtfpl.net/wp-content/uploads/2012/12/wtfpl-badge-4.png" width="80" height="15" alt="WTFPL" /></a>
