using AForge;
using AForge.Imaging;
using AForge.Imaging.Filters;
using AForge.Math.Geometry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// First attempt at making the algorithm, worked well for beginner, but too slow/unstable for intermediate/expert
// Also extremely messy, due to attempts and fails at optimizing the code for intermediate/expert
namespace Sweeper_Bot
{
    class MineSweeper
    {
        private readonly Macro macro = new Macro();
        private static List<SweeperGrid> grid = new List<SweeperGrid>();
        private static List<int> available = new List<int>();
        private static List<int> bombCoord = new List<int>();
        private static List<int> hiddenCoord = new List<int>();
        private static System.Drawing.Point p1 = new System.Drawing.Point();
        private static bool bombDetected = false;
        private static bool guess = false;
        private static bool guessCheck = false;
        private static int squaresProcessed = 0;
        private static bool finishedProcessing = false;
        private static bool clicked = false;
        private static int rowCount = 0;
        private static List<bool> hiddenImage = GetHash(new Bitmap(Properties.Resources.hidden));
        private static List<bool> blank = GetHash(new Bitmap(Properties.Resources.blank));
        private static List<bool> bomb = GetHash(new Bitmap(Properties.Resources.bomb));
        private static List<bool> num1 = GetHash(new Bitmap(Properties.Resources._1));
        private static List<bool> num2 = GetHash(new Bitmap(Properties.Resources._2));
        private static List<bool> num3 = GetHash(new Bitmap(Properties.Resources._3));
        private static List<bool> num4 = GetHash(new Bitmap(Properties.Resources._4));
        private static List<bool> num5 = GetHash(new Bitmap(Properties.Resources._5));
        private static List<bool> num6 = GetHash(new Bitmap(Properties.Resources._6));
        //private static List<bool> num7 = GetHash(new Bitmap(Properties.Resources._7));
        //private static List<bool> num8 = GetHash(new Bitmap(Properties.Resources._8));

        public void Begin(ToolStripStatusLabel toolStripStatusLabel1)
        {
            bombDetected = false;
            finishedProcessing = false;
            guess = false;
            guessCheck = false;
            grid.Clear();
            clicked = false;
            available.Clear();
            bombCoord.Clear();
            hiddenCoord.Clear();
            rowCount = 0;
            squaresProcessed = 0;
            ProcessMove(GetHandle(), toolStripStatusLabel1);
        }

        // Process image
        private void ProcessImage(Bitmap bitmap)
        {
            // lock image
            BitmapData bitmapData = bitmap.LockBits(
                new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                ImageLockMode.ReadWrite, bitmap.PixelFormat);

            // step 1 - turn background to black
            ColorFiltering colorFilter = new ColorFiltering
            {
                Blue = new IntRange(0, 64),
                FillOutsideRange = false
            };

            colorFilter.ApplyInPlace(bitmapData);

            // step 2 - locating objects
            BlobCounter blobCounter = new BlobCounter
            {
                FilterBlobs = true,
                MinHeight = 5,
                MinWidth = 5
            };
            blobCounter.ProcessImage(bitmapData);
            Blob[] blobs = blobCounter.GetObjectsInformation();
            bitmap.UnlockBits(bitmapData);

            // step 3 - check objects' type and highlight
            SimpleShapeChecker shapeChecker = new SimpleShapeChecker();
            //Pen colorPen = new Pen(Color.Yellow, 2);   // quadrilateral with known sub-type
            using (Graphics g = Graphics.FromImage(bitmap)) // SourceImage is a Bitmap object
            {
                for (int i = 0, n = blobs.Length; i < n; i++)
                {
                    List<IntPoint> edgePoints = blobCounter.GetBlobsEdgePoints(blobs[i]);

                    // is triangle or quadrilateral
                    if (shapeChecker.IsConvexPolygon(edgePoints, out List<IntPoint> corners))
                    {
                        // get sub-type
                        PolygonSubType subType = shapeChecker.CheckPolygonSubType(corners);

                        if (subType != PolygonSubType.Unknown)
                        {
                            if (corners.Count == 4)
                            {
                                // ignore the application window itself
                                if (corners[0].X <= 15)
                                {
                                    continue;
                                }
                                else
                                {
                                    //g.DrawPolygon(colorPen, ToPointsArray(corners));
                                    int right = corners[0].X, left = corners[0].X, top = corners[0].Y, bottom = corners[0].Y;
                                    for (int j = 0; j < corners.Count; j++)
                                    {
                                        if (corners[j].X > right)
                                        {
                                            right = corners[j].X;
                                        }
                                        if (corners[j].X < left)
                                        {
                                            left = corners[j].X;
                                        }
                                        if (corners[j].Y > bottom)
                                        {
                                            bottom = corners[j].Y;
                                        }
                                        if (corners[j].Y < top)
                                        {
                                            top = corners[j].Y;
                                        }
                                    }
                                    Rectangle section = new Rectangle(new System.Drawing.Point(left, top), new Size(right - left, bottom - top));
                                    IntPoint center = new IntPoint(((right - left) / 2) + left, ((top - bottom) / 2) + top);
                                    grid.Add(DetectStatus(CropImage(bitmap, section), center, section));
                                }
                            }
                        }
                    }
                }
            }
            //colorPen.Dispose();
            //sweeperGrid[10].Save(@".\image2.png");

            // put new image to clipboard
            //bitmap.Save(@".\image.png");
        }

        // Update the grid
        private void UpdateImage(Bitmap bitmap)
        {
            for (int i = grid.Count - 1; i >= 0; i--)
            {
                SweeperGrid tmpSquare = DetectStatus(CropImage(bitmap, grid[i].Rect), grid[i].Center, grid[i].Rect);
                if (grid[i].Num == tmpSquare.Num && !tmpSquare.BombVisible)
                {
                    continue;
                }
                else
                {
                    grid[i] = tmpSquare;
                    if (!grid[i].Hidden)
                    {
                        finishedProcessing = false;
                        clicked = false;
                        //hiddenCoord.Remove(i);
                        if (grid[i].Num != 0)
                        {
                            available.Add(i);
                            //hiddenCoord.RemoveAt(i);
                        }
                        else
                        {
                            //hiddenCoord.RemoveAt(i);
                        }
                    }
                }
            }
            //Console.WriteLine(hiddenCoord.Count);
        }

        public SweeperGrid DetectStatus(Bitmap bmp, IntPoint center, Rectangle rect)
        {
            List<bool> imageInfo = GetHash(new Bitmap(bmp));

            //determine the number of equal pixel (x of 256)
            int pixelHidden = imageInfo.Zip(hiddenImage, (j, k) => j == k).Count(eq => eq);
            int pixelBlank = imageInfo.Zip(blank, (j, k) => j == k).Count(eq => eq);
            int pixelBomb = imageInfo.Zip(bomb, (j, k) => j == k).Count(eq => eq);
            int pixelNum1 = imageInfo.Zip(num1, (j, k) => j == k).Count(eq => eq);
            int pixelNum2 = imageInfo.Zip(num2, (j, k) => j == k).Count(eq => eq);
            int pixelNum3 = imageInfo.Zip(num3, (j, k) => j == k).Count(eq => eq);
            int pixelNum4 = imageInfo.Zip(num4, (j, k) => j == k).Count(eq => eq);
            int pixelNum5 = imageInfo.Zip(num5, (j, k) => j == k).Count(eq => eq);
            int pixelNum6 = imageInfo.Zip(num6, (j, k) => j == k).Count(eq => eq);
            //int pixelNum7 = imageInfo.Zip(num7, (j, k) => j == k).Count(eq => eq);
            //int pixelNum8 = imageInfo.Zip(num8, (j, k) => j == k).Count(eq => eq);
            int pixelNum7 = 0, pixelNum8 = 0;

            if (pixelHidden > pixelBlank && pixelHidden > pixelNum1 && pixelHidden > pixelNum2 && pixelHidden > pixelNum3 && pixelHidden > pixelNum4 &&
                pixelHidden > pixelNum5 && pixelHidden > pixelNum6 && pixelHidden > pixelNum7 && pixelHidden > pixelNum8 && pixelHidden > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, true, center, 10, false, false, false);
                
            }
            else if (pixelBlank > pixelHidden && pixelBlank > pixelNum1 && pixelBlank > pixelNum2 && pixelBlank > pixelNum3 && pixelBlank > pixelNum4 &&
                pixelBlank > pixelNum5 && pixelBlank > pixelNum6 && pixelBlank > pixelNum7 && pixelBlank > pixelNum8 && pixelBlank > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 0, false, false, true);
            }
            else if (pixelNum1 > pixelHidden && pixelNum1 > pixelBlank && pixelNum1 > pixelNum2 && pixelNum1 > pixelNum3 && pixelNum1 > pixelNum4 &&
                pixelNum1 > pixelNum5 && pixelNum1 > pixelNum6 && pixelNum1 > pixelNum7 && pixelNum1 > pixelNum8 && pixelNum1 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 1, false, false, false);
            }
            else if (pixelNum2 > pixelHidden && pixelNum2 > pixelBlank && pixelNum2 > pixelNum1 && pixelNum2 > pixelNum3 && pixelNum2 > pixelNum4 &&
                pixelNum2 > pixelNum5 && pixelNum2 > pixelNum6 && pixelNum2 > pixelNum7 && pixelNum2 > pixelNum8 && pixelNum2 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 2, false, false, false);
            }
            else if (pixelNum3 > pixelHidden && pixelNum3 > pixelBlank && pixelNum3 > pixelNum1 && pixelNum3 > pixelNum2 && pixelNum3 > pixelNum4 &&
                pixelNum3 > pixelNum5 && pixelNum3 > pixelNum6 && pixelNum3 > pixelNum7 && pixelNum3 > pixelNum8 && pixelNum3 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 3, false, false, false);
            }
            else if (pixelNum4 > pixelHidden && pixelNum4 > pixelBlank && pixelNum4 > pixelNum1 && pixelNum4 > pixelNum2 && pixelNum4 > pixelNum3 &&
                pixelNum4 > pixelNum5 && pixelNum4 > pixelNum6 && pixelNum4 > pixelNum7 && pixelNum4 > pixelNum8 && pixelNum4 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 4, false, false, false);
            }
            else if (pixelNum5 > pixelHidden && pixelNum5 > pixelBlank && pixelNum5 > pixelNum1 && pixelNum5 > pixelNum2 && pixelNum5 > pixelNum4 &&
                pixelNum5 > pixelNum3 && pixelNum5 > pixelNum6 && pixelNum5 > pixelNum7 && pixelNum5 > pixelNum8 && pixelNum5 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 5, false, false, false);
            }
            else if (pixelNum6 > pixelHidden && pixelNum6 > pixelBlank && pixelNum6 > pixelNum1 && pixelNum6 > pixelNum2 && pixelNum6 > pixelNum4 &&
                pixelNum6 > pixelNum5 && pixelNum6 > pixelNum3 && pixelNum6 > pixelNum7 && pixelNum6 > pixelNum8 && pixelNum6 > pixelBomb)
            {
                return new SweeperGrid(bmp, rect, false, center, 6, false, false, false);
            }
            /*else if (pixelNum7 > pixelHidden && pixelNum7 > pixelBlank && pixelNum7 > pixelNum1 && pixelNum7 > pixelNum2 && pixelNum7 > pixelNum4 &&
                pixelNum7 > pixelNum5 && pixelNum7 > pixelNum6 && pixelNum7 > pixelNum3 && pixelNum7 > pixelNum8 && pixelNum7 > pixelBomb)
            {
                SweeperGrid sweeperInfo = new SweeperGrid(bmp, rect, false, center, 7, false, false, false);
                return sweeperInfo;
            }
            else if (pixelNum8 > pixelHidden && pixelNum8 > pixelBlank && pixelNum8 > pixelNum1 && pixelNum8 > pixelNum2 && pixelNum8 > pixelNum4 &&
                pixelNum8 > pixelNum5 && pixelNum8 > pixelNum6 && pixelNum8 > pixelNum7 && pixelNum8 > pixelNum3 && pixelNum8 > pixelBomb)
            {
                SweeperGrid sweeperInfo = new SweeperGrid(bmp, rect, false, center, 8, false, false, false);
                return sweeperInfo;
            }*/
            else if (pixelBomb > pixelHidden && pixelBomb > pixelBlank && pixelBomb > pixelNum1 && pixelBomb > pixelNum2 && pixelBomb > pixelNum4 &&
                pixelBomb > pixelNum5 && pixelBomb > pixelNum6 && pixelBomb > pixelNum7 && pixelBomb > pixelNum3 && pixelBomb > pixelNum8)
            {
                return new SweeperGrid(bmp, rect, false, center, 9, true, true, true);
            }
            else
            {
                return new SweeperGrid(bmp, rect, true, center, 10, false, false, false);
            }
        }

        public static List<bool> GetHash(Bitmap bmpSource)
        {
            List<bool> lResult = new List<bool>();
            //create new image with 16x16 pixel
            Bitmap bmpMin = new Bitmap(bmpSource, new Size(16, 16));
            for (int j = 0; j < bmpMin.Height; j++)
            {
                for (int i = 0; i < bmpMin.Width; i++)
                {
                    //reduce colors to true / false                
                    lResult.Add(bmpMin.GetPixel(i, j).GetBrightness() < 0.5f);
                }
            }
            return lResult;
        }

        // Conver list of AForge.NET's points to array of .NET points
        private System.Drawing.Point[] ToPointsArray(List<IntPoint> points)
        {
            System.Drawing.Point[] array = new System.Drawing.Point[points.Count];

            for (int i = 0, n = points.Count; i < n; i++)
            {
                array[i] = new System.Drawing.Point(points[i].X, points[i].Y);
            }

            return array;
        }

        public Bitmap CropImage(Bitmap source, Rectangle section)
        {
            // An empty bitmap which will hold the cropped image
            Bitmap bmp = new Bitmap(section.Width, section.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Draw the given area (section) of the source image
                // at location 0,0 on the empty bitmap (bmp)
                g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);
            }
            return RemoveColor(bmp);
        }

        // remove color for better detection (probably don't need but beneficial to try out)
        public Bitmap RemoveColor(Bitmap source)
        {
            using (Graphics gr = Graphics.FromImage(source)) // SourceImage is a Bitmap object
            {
                var gray_matrix = new float[][] {
                new float[] { 0.299f, 0.299f, 0.299f, 0, 0 },
                new float[] { 0.587f, 0.587f, 0.587f, 0, 0 },
                new float[] { 0.114f, 0.114f, 0.114f, 0, 0 },
                new float[] { 0,      0,      0,      1, 0 },
                new float[] { 0,      0,      0,      0, 1 }
            };

                var ia = new System.Drawing.Imaging.ImageAttributes();
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix(gray_matrix));
                ia.SetThreshold((float)0.7); // Change this threshold as needed
                var rc = new Rectangle(0, 0, source.Width, source.Height);
                gr.DrawImage(source, rc, 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, ia);
            }
            return source;
        }

        // use quicksort to ensure grid in the proper order
        public void OrganizeGrid()
        {
            QuickSortY(0, grid.Count - 1);
            int coordY = grid[0].Center.Y;
            int left = 0;
            for (int i = 0; i < grid.Count; i++)
            {
                if (Math.Abs(grid[i].Center.Y - coordY) > 10)
                {
                    if (rowCount == 0)
                    {
                        rowCount = i;
                    }
                    //Console.WriteLine("Sorting at: " + i);
                    QuickSortX(left, i - 1);
                    coordY = grid[i].Center.Y;
                    left = i;
                }
                // the algorithm seems to miss the last row, this forces it to sort the last row
                if (i == (grid.Count - 1))
                {
                    QuickSortX(left, i);
                }
            }
        }

        private static void QuickSortY(int left, int right)
        {
            if (left < right)
            {
                int pivot = PartitionY(left, right);

                if (pivot > 1)
                {
                    QuickSortY(left, pivot - 1);
                }
                if (pivot + 1 < right)
                {
                    QuickSortY(pivot + 1, right);
                }
            }
        }

        private static int PartitionY(int left, int right)
        {
            int pivot = grid[left].Center.Y;
            while (true)
            {

                while (grid[left].Center.Y < pivot)
                {
                    left++;
                }

                while (grid[right].Center.Y > pivot)
                {
                    right--;
                }

                if (left < right)
                {
                    if (grid[left].Center.Y == grid[right].Center.Y) return right;

                    SweeperGrid temp = grid[left];
                    grid[left] = grid[right];
                    grid[right] = temp;
                }
                else
                {
                    return right;
                }
            }
        }

        private static void QuickSortX(int left, int right)
        {
            if (left < right)
            {
                int pivot = PartitionX(left, right);

                if (pivot > 1)
                {
                    QuickSortX(left, pivot - 1);
                }
                if (pivot + 1 < right)
                {
                    QuickSortX(pivot + 1, right);
                }
            }

        }

        private static int PartitionX(int left, int right)
        {
            int pivot = grid[left].Center.X;
            while (true)
            {

                while (grid[left].Center.X < pivot)
                {
                    left++;
                }

                while (grid[right].Center.X > pivot)
                {
                    right--;
                }

                if (left < right)
                {
                    if (grid[left].Center.X == grid[right].Center.X) return right;

                    SweeperGrid temp = grid[left];
                    grid[left] = grid[right];
                    grid[right] = temp;


                }
                else
                {
                    return right;
                }
            }
        }

        private static void InsertionSort()
        {
            for (int i = 0; i < available.Count - 1; i++)
            {
                for (int j = i + 1; j > 0; j--)
                {
                    if (grid[available[j - 1]].Num > grid[available[j]].Num)
                    {
                        int temp = available[j - 1];
                        available[j - 1] = available[j];
                        available[j] = temp;
                    }
                }
            }
        }

        public void PrintGrid()
        {
            int coordY = grid[0].Center.Y;
            for (int i = 0; i < grid.Count; i++)
            {
                if ((grid[i].Center.Y - coordY) > 10)
                {
                    Console.WriteLine("");
                    coordY = grid[i].Center.Y;
                }
                if (grid[i].Num < 9)
                {
                    if (grid[i].Processed)
                    {
                        Console.Write("{" + grid[i].Num + "}");
                    }
                    else
                    {
                        Console.Write("[" + grid[i].Num + "]");
                    }
                }
                else if (grid[i].Num == 9)
                {
                    if (grid[i].Processed)
                    {
                        Console.Write("{" + "X" + "}");
                    }
                    else
                    {
                        Console.Write("[" + "X" + "]");
                    }
                }
                else
                {
                    if (grid[i].Processed)
                    {
                        Console.Write("{" + "~" + "}");
                    }
                    else
                    {
                        Console.Write("[" + "~" + "]");
                    }
                }
            }
            Console.WriteLine("");
            Console.WriteLine("---------------------------");
        }

        public void SaveGrid()
        {
            if (!Directory.Exists(@".\grid"))
            {
                Directory.CreateDirectory(@".\grid");
            }
            for (int i = 0; i < grid.Count; i++)
            {
                grid[i].BMP.Save(@".\grid\grid" + (i + 1) + ".png");
            }
        }

        public bool CheckForBomb()
        {
            for (int i = 0; i < grid.Count; i++)
            {
                if (grid[i].BombVisible)
                {
                    return true;
                }
            }
            return false;
        }

        // the main algorithm
        public void FindBestMove(ToolStripStatusLabel toolStripStatusLabel1)
        {
            int squareToProcess = 0;
            if (available.Count == 0)
            {
                for (int i = 0; i < grid.Count; i++)
                {
                    if (!grid[i].Hidden && grid[i].Num != 0)
                    {
                        available.Add(i);
                    }
                    else if (grid[i].Hidden)
                    {
                        if (hiddenCoord.Any(s => i == s))
                        {
                            continue;
                        }
                        else
                        {
                            hiddenCoord.Add(i);
                        }
                    }
                }
                if (available.Count == 0)
                {
                    squareToProcess = (grid.Count / 2) + 1;
                    int clickCenter = (((grid.Count / rowCount) / 2) * rowCount) - (rowCount / 2);
                    p1.X = grid[clickCenter].Center.X;
                    p1.Y = grid[clickCenter].Center.Y;
                    clicked = true;
                }
            }
            else
            {
                InsertionSort();
                for (int i = 0; i < available.Count; i++)
                {
                    int coordY = grid[available[i]].Center.Y;
                    int possibleBomb = 0;

                    String missingSide = "";
                    List<int> surroundingCoord = new List<int>();
                    // check left
                    if ((available[i] - 1) >= 0)
                    {
                        if (Math.Abs(coordY - grid[available[i] - 1].Center.Y) < 10)
                        {
                            if (grid[available[i] - 1].Hidden)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] - 1);
                            }
                        }
                    }
                    else
                    {
                        missingSide += "Left ";
                    }
                    // check right
                    if ((available[i] + 1) < grid.Count)
                    {
                        if (Math.Abs(coordY - grid[available[i] + 1].Center.Y) < 10)
                        {
                            if (grid[available[i] + 1].Hidden)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] + 1);
                            }
                        }
                    }
                    else
                    {
                        missingSide += "Right ";
                    }
                    // check top
                    if ((available[i] - rowCount) >= 0)
                    {
                        if (grid[available[i] - rowCount].Hidden)
                        {
                            possibleBomb++;
                            surroundingCoord.Add(available[i] - rowCount);
                        }
                    }
                    else
                    {
                        missingSide += "Top ";
                    }
                    // check bottom
                    if ((available[i] + rowCount) < grid.Count)
                    {
                        if (grid[available[i] + rowCount].Hidden)
                        {
                            possibleBomb++;
                            surroundingCoord.Add(available[i] + rowCount);
                        }
                    }
                    else
                    {
                        missingSide += "Bottom ";
                    }
                    // check top left
                    if ((available[i] - rowCount) >= 0 && (available[i] - rowCount - 1) >= 0)
                    {
                        if (grid[available[i] - rowCount - 1].Hidden)
                        {
                            if (Math.Abs(grid[available[i] - rowCount].Center.Y - grid[available[i] - rowCount - 1].Center.Y) < 10)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] - rowCount - 1);
                            }
                        }
                    }
                    // check top right
                    if ((available[i] - rowCount) >= 0 && (available[i] - rowCount + 1) >= 0)
                    {
                        if (grid[available[i] - rowCount + 1].Hidden)
                        {
                            if (Math.Abs(grid[available[i] - rowCount].Center.Y - grid[available[i] - rowCount + 1].Center.Y) < 10)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] - rowCount + 1);
                            }
                        }
                    }
                    // check bottom left
                    if ((available[i] + rowCount) < grid.Count && (available[i] + rowCount - 1) < grid.Count)
                    {
                        if (grid[available[i] + rowCount - 1].Hidden)
                        {
                            if (Math.Abs(grid[available[i] + rowCount].Center.Y - grid[available[i] + rowCount - 1].Center.Y) < 10)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] + rowCount - 1);
                            }
                        }
                    }
                    // check bottom right
                    if ((available[i] + rowCount) < grid.Count && (available[i] + rowCount + 1) < grid.Count)
                    {
                        if (grid[available[i] + rowCount + 1].Hidden)
                        {
                            if (Math.Abs(grid[available[i] + rowCount].Center.Y - grid[available[i] + rowCount + 1].Center.Y) < 10)
                            {
                                possibleBomb++;
                                surroundingCoord.Add(available[i] + rowCount + 1);
                            }
                        }
                    }
                    // determine if bomb or not: click if it's safe, mark if it's dangerous, pass if undetermined
                    int confirmedBomb = 0;
                    for (int j = surroundingCoord.Count - 1; j >= 0; j--)
                    {
                        if (bombCoord.Any(s => surroundingCoord[j].Equals(s)) || grid[surroundingCoord[j]].IsBomb)
                        {
                            confirmedBomb++;
                            surroundingCoord.RemoveAt(j);
                        }
                    }
                    // the guessing algorithm, slightly random but should be wary of bombs
                    if (guess)
                    {
                        //method 1
                        if (surroundingCoord.Count > 0)
                        {
                            for (int j = 0; j < grid.Count; j++)
                            {
                                if (grid[i].Hidden)
                                {
                                    if (bombCoord.Any(s => i.Equals(s)) || grid[i].IsBomb)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        //Console.WriteLine("guess method 1 failing");
                                        squareToProcess = (surroundingCoord[j]) + 1;
                                        guessCheck = true;
                                        squaresProcessed = 0;
                                        finishedProcessing = false;
                                        guess = false;
                                        p1.X = grid[surroundingCoord[j]].Center.X;
                                        p1.Y = grid[surroundingCoord[j]].Center.Y;
                                        clicked = true;
                                        break;
                                    }
                                }
                            }
                        }
                        if (guessCheck)
                        {
                            break;
                        }
                        else
                        {
                            // method 2
                            for (int j = 0; j < surroundingCoord.Count; j++)
                            {
                                if (bombCoord.Any(s => surroundingCoord[j].Equals(s)) || grid[surroundingCoord[j]].IsBomb)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (grid[surroundingCoord[j]].Num == 10)
                                    {
                                        //Console.WriteLine("guess method 2 failing");
                                        squareToProcess = (surroundingCoord[j]) + 1;
                                        guessCheck = true;
                                        guess = false;
                                        finishedProcessing = false;
                                        squaresProcessed = 0;
                                        p1.X = grid[surroundingCoord[j]].Center.X;
                                        p1.Y = grid[surroundingCoord[j]].Center.Y;
                                        clicked = true;
                                        break;
                                    }
                                }
                            }
                        }
                        if (guessCheck)
                        {
                            break;
                        }
                        else
                        {
                            squaresProcessed = grid.Count * 2;
                        }
                    }
                    //Console.WriteLine("Processing: " + grid[available[i]].Num + " at grid: " + (available[i] + 1) + " without side: " + missingSide + " and found possible: " + confirmedBomb + "/" + possibleBomb + " and counter is at " + surroundingCoord.Count);
                    if (possibleBomb == grid[available[i]].Num && !finishedProcessing)
                    {
                        for (int j = 0; j < surroundingCoord.Count; j++)
                        {
                            //if (!grid[surroundingCoord[j]].IsBomb)
                            //{
                                grid[surroundingCoord[j]] = new SweeperGrid(grid[surroundingCoord[j]].BMP, grid[surroundingCoord[j]].Rect, true, grid[surroundingCoord[j]].Center, 9, false, true, true);
                                if (bombCoord.Any(s => surroundingCoord[j].Equals(s)))
                                {
                                    continue;
                                }
                                else
                                {
                                    bombCoord.Add(surroundingCoord[j]);
                                }
                                //Console.WriteLine("Grid " + (available[i] + 1) + " has a bomb nearby at " + (surroundingCoord[j] + 1));
                            //}
                        }
                        grid[available[i]] = new SweeperGrid(grid[available[i]].BMP, grid[available[i]].Rect, false, grid[available[i]].Center, grid[available[i]].Num, false, false, true);
                        squareToProcess = available[i] + 1;
                    }
                    else if (possibleBomb > grid[available[i]].Num && finishedProcessing)
                    {
                        if (confirmedBomb == grid[available[i]].Num)
                        {
                            //Console.WriteLine(grid[available[i]].Num + " is at Grid: " + (available[i] + 1) + " has " + confirmedBomb + " of " + possibleBomb);
                            for (int j = 0; j < surroundingCoord.Count; j++)
                            {
                                if (bombCoord.Any(s => surroundingCoord[j].Equals(s)) || grid[surroundingCoord[j]].IsBomb)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (grid[surroundingCoord[j]].Num == 10)
                                    {
                                        //Console.WriteLine("if statement failing");
                                        squareToProcess = (surroundingCoord[j]) + 1;
                                        p1.X = grid[surroundingCoord[j]].Center.X;
                                        p1.Y = grid[surroundingCoord[j]].Center.Y;
                                        clicked = true;
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                    if (!finishedProcessing && i == available.Count - 1)
                    {
                        finishedProcessing = true;
                        //Console.WriteLine("done processing");
                    }
                    else if (finishedProcessing && i == available.Count - 1 && !guess)
                    {
                        //Console.WriteLine("should start guessing");
                        guess = true;
                        squaresProcessed = grid.Count;
                        break;
                    }
                    //Console.WriteLine(grid[available[i]].Num + " is at Grid: " + (available[i] + 1) + " has " + confirmedBomb + " of " + possibleBomb + " | finished processing: " + finishedProcessing);

                    //Console.WriteLine(available[i]+1 + ":" + grid[60].IsBomb);
                }
                squaresProcessed++;
                if (!guess)
                {
                    toolStripStatusLabel1.Text = "Processing node: " + squareToProcess;
                }
                else
                {
                    toolStripStatusLabel1.Text = "Guessing next move";
                }
                //Console.WriteLine("Processing: " + squaresProcessed + " | " + "Out of: " + available.Count + " nodes");
            }
        }

        public IntPtr GetHandle()
        {
            Process[] processlist = Process.GetProcesses();
            String[] process = null;
            process = new string[200];
            int i = 0;
            IntPtr handle = IntPtr.Zero;
            foreach (Process theprocess in processlist)
            {
                process[i] = theprocess.MainWindowTitle;
                if (process[i].Equals("Minesweeper"))
                {
                    handle = theprocess.MainWindowHandle;
                    break;
                }
                i++;
            }
            return handle;
        }

        public void ProcessMove(IntPtr handle, ToolStripStatusLabel toolStripStatusLabel1)
        {
            if (handle == IntPtr.Zero)
            {
                toolStripStatusLabel1.Text = "Process not found";
                return;
            }
            toolStripStatusLabel1.Text = "Process found";
            ProcessImage(this.macro.GetWindowBitmap(handle));
            OrganizeGrid();
            //SaveGrid();
            //PrintGrid();
            FindBestMove(toolStripStatusLabel1);
            while (squaresProcessed < (grid.Count * 2))
            {
                this.macro.ClickOnPoint(handle, p1);
                if (clicked)
                {
                    UpdateImage(this.macro.GetWindowBitmap(handle));
                    if (squaresProcessed > grid.Count)
                    {
                        int bomb = 0;
                        int hidden = 0;
                        for (int i = 0; i < grid.Count; i++)
                        {
                            if (bombCoord.Any(s => i.Equals(s)) || grid[i].IsBomb)
                            {
                                bomb++;
                            }
                            if (grid[i].Hidden)
                            {
                                hidden++;
                            }
                        }
                        //Console.WriteLine(hidden + " | " + bomb);
                        if (hidden == bomb || hidden == bombCoord.Count)
                        {
                            break;
                        }
                    }
                }
                if (CheckForBomb())
                {
                    //Console.WriteLine("Detected bomb");
                    bombDetected = CheckForBomb();
                    break;
                }
                FindBestMove(toolStripStatusLabel1);
                //PrintGrid();
            }
            if (bombDetected && guessCheck)
            {
                toolStripStatusLabel1.Text = "Guessed incorrectly";
            }
            else if (bombDetected)
            {
                toolStripStatusLabel1.Text = "Failed to solve";
            }
            else
            {
                int bomb = 0;
                int hidden = 0;
                for (int i = 0; i < grid.Count; i++)
                {
                    if (bombCoord.Any(s => i.Equals(s)) || grid[i].IsBomb)
                    {
                        bomb++;
                    }
                    if (grid[i].Hidden)
                    {
                        hidden++;
                    }
                }
                //Console.WriteLine(hidden + " | " + bomb);
                if (hidden == bomb || hidden == bombCoord.Count)
                {
                    toolStripStatusLabel1.Text = "Successfully solved";
                }
                else
                {
                    toolStripStatusLabel1.Text = "Failed to guess";
                }
            }
        }

        public struct SweeperGrid
        {
            private Bitmap _bmp;
            private Rectangle _rect;
            private bool _hidden;
            private IntPoint _center;
            private int _num;
            private bool _bombvisible;
            private bool _isbomb;
            private bool _processed;

            public SweeperGrid(Bitmap bmp, Rectangle rect, bool hidden, IntPoint center, int num, bool bombvisible, bool isbomb, bool processed)
            {
                _bmp = bmp;
                _rect = rect;
                _hidden = hidden;
                _center = center;
                _num = num;
                _bombvisible = bombvisible;
                _isbomb = isbomb;
                _processed = processed;
            }

            public Bitmap BMP
            {
                get { return _bmp; }
                set { _bmp = value; }
            }

            public Rectangle Rect
            {
                get { return _rect; }
                set { _rect = value; }
            }

            public bool Hidden
            {
                get { return _hidden; }
                set { _hidden = value; }
            }

            public IntPoint Center
            {
                get { return _center; }
                set { _center = value; }
            }

            public int Num
            {
                get { return _num; }
                set { _num = value; }
            }

            public bool BombVisible
            {
                get { return _bombvisible; }
                set { _bombvisible = value; }
            }

            public bool IsBomb
            {
                get { return _isbomb; }
                set { _isbomb = value; }
            }

            public bool Processed
            {
                get { return _processed; }
                set { _processed = value; }
            }
        }
    }
}
