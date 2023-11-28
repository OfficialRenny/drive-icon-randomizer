using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Win32;

namespace DriveIconRandomizer;

internal class Program
{
    private static string _workingDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
               "DriveIconRandomizer");

    private static void Main(string[] args)
    {
        if (args.Length < 1 || args[0] == "--help")
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("DriveIconRandomizer.exe <path to image directory>");
            Console.WriteLine("\tSets the drive icons for all fixed drives to random images in the specified directory");
            Console.WriteLine("DriveIconRandomizer.exe --clear");
            Console.WriteLine("\tClears all drive icons");
            return;
        }

        if (!Directory.Exists(_workingDir))
            Directory.CreateDirectory(_workingDir);

        var mappedDrives = DriveInfo.GetDrives().Where(x => x.DriveType == DriveType.Fixed).ToList();

        if (args[0] == "--clear")
        {
            foreach (var drive in mappedDrives)
                try
                {
                    var driveLetter = drive.Name[..1];
                    ClearDriveIcon(driveLetter);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error clearing icon for drive " + drive.Name + ": " + ex.Message);
                }

            Directory.GetFiles(_workingDir, "*.ico").ToList().ForEach(File.Delete);
            return;
        }

        var imageDir = args[0];
        var driveIconDir = new DirectoryInfo(imageDir);

        if (!driveIconDir.Exists)
        {
            Console.WriteLine("Directory does not exist: " + imageDir);
            return;
        }

        var fileExtensions = new string[]
        {
            ".jpg",
            ".jpeg",
            ".png",
            ".gif",
            ".bmp"
        };

        var allImages = driveIconDir.GetFiles("*.*", SearchOption.AllDirectories).Where(x =>
            fileExtensions.Contains(x.Extension.ToLower())).ToList();

        var random = new Random();
        var randomImages = allImages.OrderBy(_ => random.Next()).Take(mappedDrives.Count).ToList();

        for (var i = 0; i < mappedDrives.Count; i++)
        {
            var drive = mappedDrives[i];
            var driveLetter = drive.Name[..1];
            var image = randomImages[i];

#if DEBUG
            Console.WriteLine("Setting icon for drive " + drive.Name + " to " + image.Name);
#endif

            var icon = IconFromImage(Image.FromFile(image.FullName));
            var iconPath = Path.Combine(_workingDir, $"{image.Name}.ico");

            try
            {
                using var fs = new FileStream(iconPath, FileMode.Create);
                icon.Save(fs);

                SetDriveIcon(driveLetter, iconPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error setting icon for drive " + drive.Name + ": " + ex.Message);
            }
        }
    }

    private static Icon IconFromImage(Image img)
    {
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        // Header
        bw.Write((short)0); // 0 : reserved
        bw.Write((short)1); // 2 : 1=ico, 2=cur
        bw.Write((short)1); // 4 : number of images
        // Image directory
        var w = img.Width;
        if (w >= 256) w = 0;
        bw.Write((byte)w); // 0 : width of image
        var h = img.Height;
        if (h >= 256) h = 0;
        bw.Write((byte)h); // 1 : height of image
        bw.Write((byte)0); // 2 : number of colors in palette
        bw.Write((byte)0); // 3 : reserved
        bw.Write((short)0); // 4 : number of color planes
        bw.Write((short)0); // 6 : bits per pixel
        var sizeHere = ms.Position;
        bw.Write(0); // 8 : image size
        var start = (int)ms.Position + 4;
        bw.Write(start); // 12: offset of image data
        // Image data
        img.Save(ms, ImageFormat.Png);
        var imageSize = (int)ms.Position - start;
        ms.Seek(sizeHere, SeekOrigin.Begin);
        bw.Write(imageSize);
        ms.Seek(0, SeekOrigin.Begin);

        // And load it
        return new Icon(ms);
    }

    private static void SetDriveIcon(string driveLetter, string iconPath)
    {
        var key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\" +
                                                    driveLetter + @"\DefaultIcon");
        key.SetValue("", iconPath);

        key.Close();
    }

    private static void ClearDriveIcon(string driveLetter)
    {
        // clear the drive icon for a given drive letter
        var key = Registry.LocalMachine.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\" +
                                                    driveLetter + @"\DefaultIcon");
        key.DeleteValue("");
        key.Close();
    }
}