<%@Page Language="C#"%>
<%@ Import Namespace="System" %>  
<%@ Import Namespace="System.Text" %>  
<%@ Import Namespace="System.IO" %>   
<%@ Import Namespace="System.IO.Compression" %> 


<script runat="server">
   void Page_Load(Object sender, EventArgs e)
   {
  
   String sourceDir = Request["sourceDir"];
   String sourceZip = Request["sourceZip"];
 
   string startPath = @sourceDir;
   string zipPath   = @sourceZip;

   ZipFile.CreateFromDirectory(startPath, zipPath, CompressionLevel.Fastest, true,Encoding.UTF8);

   Response.Write("OK");
}
</script>