using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

//參考文章:https://heavenchou.buddhason.org/node/410
//參考文章:https://heavenchou.buddhason.org/node/411
//參考文章:https://vito-note.blogspot.com/2015/04/google-drive-api.html

class Program
{
    static void Main()
    {
        Program program = new Program();
        //取得檔案
        program.getFile();

        //上傳檔案
        //Console.WriteLine("【上傳檔案】");
        //Console.WriteLine("請輸入上傳檔案路徑:");
        //var filePath = Console.ReadLine();
        //Console.WriteLine("請輸入資料夾ID(無請輸入空白):");
        //var uploadParentFolderId = Console.ReadLine();
        //program.Upload(filePath, uploadParentFolderId);

        //更新檔案
        //Console.WriteLine("【更新檔案】");
        //Console.WriteLine("請輸入更新檔案路徑:");
        //var updateFilePath = Console.ReadLine();
        //Console.WriteLine("請輸入更新檔案ID:");
        //var updateFileId = Console.ReadLine();
        //program.UpdateFile(updateFilePath, updateFileId);

        //下載檔案
        //Console.WriteLine("【下載檔案】");
        //Console.WriteLine("請輸入下載檔案ID:");
        //var fileID = Console.ReadLine();
        //Console.WriteLine("請輸入存檔的路徑:");
        //var newFilePath = Console.ReadLine();
        //program.DownloadFile(newFilePath, fileID);

        //建立目錄
        //Console.WriteLine("【建立目錄】");
        //Console.WriteLine("請輸入資料夾名稱:");
        //var folderName = Console.ReadLine();
        //Console.WriteLine("請輸入資料夾ID(無請輸入空白):");
        //var addParentFolderId = Console.ReadLine();
        //program.CreateFolderInDrive(folderName, addParentFolderId);

        //更改檔名
        //Console.WriteLine("【更改檔名】");
        //Console.WriteLine("請輸入更改檔案的ID:");
        //var folderId = Console.ReadLine();
        //Console.WriteLine("請輸入更改檔案的名稱:");
        //var folderName = Console.ReadLine();
        //program.UpdateName(folderId, folderName);

        //移動檔案
        //Console.WriteLine("【移動檔案】");
        //Console.WriteLine("請輸入移動檔案的ID:");
        //var folderId = Console.ReadLine();
        //Console.WriteLine("請輸入移動至資料夾的ID:");
        //var moveParentFolderId = Console.ReadLine();
        //program.MoveFile(folderId, moveParentFolderId);

        //刪除檔案
        //Console.WriteLine("【刪除檔案】");
        //Console.WriteLine("請輸入刪除檔案的ID:");
        //var deleteFileId = Console.ReadLine();
        //program.DeleteFile(deleteFileId);
    }

    string googleClinetId = "GoogleApisClinetId";
    string googleClinetSecret = "GoogleApisClinetSecret";

    /// <summary>
    /// 登入
    /// </summary>
    /// <param name="_googleClinetId"></param>
    /// <param name="_googleClinetSecret"></param>
    /// <returns></returns>
    private UserCredential Login(string _googleClinetId, string _googleClinetSecret)
    {
        ClientSecrets secrets = new ClientSecrets()
        {
            ClientId = _googleClinetId,
            ClientSecret = _googleClinetSecret
        };

        return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets,
            new[] { DriveService.Scope.Drive },
            "user",
            CancellationToken.None).Result;
    }

    /// <summary>
    /// 取得 Google 服務
    /// </summary>
    /// <returns></returns>
    private DriveService GetDriveService()
    {
        UserCredential credential = Login(googleClinetId, googleClinetSecret);

        // 創建 DriveService 實例
        return new DriveService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "Drive API Sample",
        });
    }

    ///取得檔案
    private void getFile()
    {
        UserCredential credential = Login(googleClinetId, googleClinetSecret);
        using (var driveService = new DriveService(new BaseClientService.Initializer() { HttpClientInitializer = credential }))
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.Q = "";//列出全部目錄檔案
            FileList listFileRequest = listRequest.Execute();

            foreach (var file in listFileRequest.Files)
            {
                Console.WriteLine(file.Parents);
                Console.WriteLine(file.Name + "(" + file.Id + ")");
            }
        }
        Console.ReadLine();
    }

    /// <summary>
    /// 上傳檔案
    /// </summary>
    /// <param name="filePath">local 端的檔案路徑</param>
    /// <param name="parentFolderId">要上傳到 google drive 上的 folder id</param>
    /// <returns></returns>
    private void Upload(string filePath, string parentFolderId = "")
    {
        DriveService service = GetDriveService();
        // 如果沒有指定父目錄，則用 root 為根目錄

        if (parentFolderId == "")
        {
            parentFolderId = "root";
        }
        // 創建 File 實例
        var fileMetadata = new Google.Apis.Drive.v3.Data.File()
        {
            //Name = Path.GetFileName(filePath)
            Name = "20240129.txt",
            Parents = new List<string> { parentFolderId }
        };

        using (var stream = new FileStream(filePath, FileMode.Open))
        {
            var request = service.Files.Create(fileMetadata, stream, GetMimeType(filePath));
            request.Fields = "id";

            try
            {
                request.Upload();
                var responseFile = request.ResponseBody;
                Console.WriteLine("File ID: " + responseFile.Id);
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
        }
    }

    /// <summary>
    /// 更新檔案
    /// </summary>
    /// <param name="updateFile"></param>
    /// <param name="fileId"></param>
    /// <returns></returns>
    private void UpdateFile(string updateFile, string fileId)
    {
        Program program = new Program();
        DriveService service = program.GetDriveService();

        var fileMetadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = Path.GetFileName(updateFile)
        };

        using (var stream = new FileStream(updateFile, FileMode.Open))
        {
            var request = service.Files.Update(fileMetadata, fileId, stream, program.GetMimeType(updateFile));
            request.Fields = "id";
            request.Upload();
        }
    }

    /// <summary>
    /// 判斷檔案類型
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    private string GetMimeType(string filePath)
    {
        string mimeType = "application/unknown";
        string ext = Path.GetExtension(filePath).ToLower();
        Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
        if (regKey != null && regKey.GetValue("Content Type") != null)
        {
            mimeType = regKey.GetValue("Content Type").ToString();
        }
        return mimeType;
    }

    /// <summary>
    /// 下載檔案
    /// </summary>
    /// <param name="newFilePath"></param>
    /// <param name="fileId"></param>
    private void DownloadFile(string newFilePath, string fileId)
    {
        DriveService service = GetDriveService();

        var request = service.Files.Get(fileId);
        using (var stream = new FileStream(newFilePath, FileMode.Create))
        {
            request.Download(stream);
        }
    }

    /// <summary>
    /// 建立目錄
    /// </summary>
    /// <param name="folderName"></param>
    /// <param name="parentFolderId"></param>
    private void CreateFolderInDrive(string folderName, string parentFolderId = "")
    {
        DriveService service = GetDriveService();

        // 如果沒有指定父目錄，則用 root 為根目錄
        if (parentFolderId == "")
        {
            parentFolderId = "root";
        }

        // 創建 File 實例
        var fileMetadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = folderName,
            MimeType = "application/vnd.google-apps.folder",
            Parents = new List<string> { parentFolderId }
        };

        // 建立目錄
        var request = service.Files.Create(fileMetadata);
        request.Fields = "id";          // 僅取回 id 屬性，亦可忽略不寫
        var folder = request.Execute();

        // 打印新子目錄的ID
        if (folder != null)
        {
            Console.WriteLine($"建立目錄 ID：{folder.Id}");
        }
        else
        {
            Console.WriteLine("建立目錄失敗。");
        }
    }

    /// <summary>
    /// 目錄或檔案更名
    /// </summary>
    /// <param name="fileId"></param>
    /// <param name="newFileName"></param>
    private void UpdateName(string fileId, string newFileName)
    {
        DriveService service = GetDriveService();

        var file = new Google.Apis.Drive.v3.Data.File();
        file.Name = newFileName;

        service.Files.Update(file, fileId).Execute();
        Console.WriteLine("OK");
    }

    /// <summary>
    /// 目錄檔案移動
    /// </summary>
    /// <param name="fileId"></param>
    /// <param name="parentFolderId"></param>
    private void MoveFile(string fileId, string parentFolderId)
    {
        DriveService service = GetDriveService();

        // 若目的目錄空白，就移到根目錄
        if (parentFolderId == "")
        {
            parentFolderId = "root";
        }

        var file = new Google.Apis.Drive.v3.Data.File();
        FilesResource.UpdateRequest request = service.Files.Update(file, fileId);

        // Update 時，不能直接修改 Parent 屬性，要用 AddParents
        request.AddParents = parentFolderId;
        request.Execute();
        Console.WriteLine("OK");
    }

    /// <summary>
    /// 刪除目錄或檔案(永久刪除，小心操作)
    /// </summary>
    /// <param name="fileId"></param>
    private void DeleteFile(string fileId)
    {
        DriveService service = GetDriveService();
        service.Files.Delete(fileId).Execute();
        Console.WriteLine("OK");
    }

}
