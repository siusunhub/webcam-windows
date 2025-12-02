using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using DirectShowLib;
using Size = OpenCvSharp.Size;

namespace WebcamWindows
{
    // Structure for video format information
    [StructLayout(LayoutKind.Sequential)]
    public struct VideoFormat
    {
        public int Width;
        public int Height;
        public string PixelFormat;
        public double FrameRate;
        
        public override string ToString()
        {
            return $"{Width}x{Height} @ {FrameRate:F1}fps ({PixelFormat})";
        }
    }

    public partial class MainForm : Form
    {
        // Core components
        private VideoCapture videoCapture;
        private Timer frameTimer;
        private Timer clockTimer;
        private Timer autoResizeTimer;
        
        // Webcam management
        private List<string> availableWebcams;
        private List<Size> supportedResolutions = new List<Size>();
        private Size currentResolution = new Size(640, 480);
        private int currentDeviceIndex = 0;
        private string currentWebcamName = "";
        private HashSet<int> failedWebcamIndices = new HashSet<int>();
        
        // Window state management
        private bool isWebcamEnabled = true;
        private bool isBorderless = false;
        private bool isFullScreen = false;
        private bool wasDisabledByClose = false;
        private FormBorderStyle originalBorderStyle;
        private FormWindowState originalWindowState;
        private Rectangle originalBounds;
        
        // Auto-resize functionality
        private bool isAutoResizing = false;
        private System.Drawing.Point preResizeLocation;
        private bool isInResizeSequence = false;
        
        // Application state
        private bool isDisposing = false;
        private readonly object lockObject = new object();
        private bool minimizeToTrayEnabled = false;

        public MainForm()
        {
            InitializeComponent();
            
            // Store original window properties
            this.originalBorderStyle = this.FormBorderStyle;
            this.originalWindowState = this.WindowState;
            this.originalBounds = this.Bounds;
            
            // Center the window on screen at startup
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // Initialize timers
            InitializeTimers();
            
            // Initialize webcams after form is created
            this.Load += MainForm_Load;
        }

        private void InitializeTimers()
        {
            // Frame processing timer (~30 FPS)
            frameTimer = new Timer();
            frameTimer.Interval = 33;
            frameTimer.Tick += FrameTimer_Tick;

            // Clock timer for window title updates
            clockTimer = new Timer();
            clockTimer.Interval = 500;
            clockTimer.Tick += ClockTimer_Tick;
            clockTimer.Start();
            
            // Auto-resize timer (5 seconds after resize)
            autoResizeTimer = new Timer();
            autoResizeTimer.Interval = 5000;
            autoResizeTimer.Tick += AutoResizeTimer_Tick;
        }

        private async void MainForm_Load(object sender, EventArgs e)
        {
            await InitializeWebcamAsync();
            SetupTrayIcon();
        }

        private async Task InitializeWebcamAsync()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                ShowSwitchingScreen("INITIALIZING...");
                
                // Discover available webcams
                await Task.Run(() => {
                    availableWebcams = GetWebcamNames();
                });
                
                if (availableWebcams.Count > 0)
                {
                    StartWebcam(0); // Start with first available webcam
                }
                else
                {
                    ShowDisabledScreen();
                    MessageBox.Show("No webcam devices found!\n\nTip: Try 'Refresh Webcams' from the context menu.", 
                                   "No Webcams", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowDisabledScreen();
                MessageBox.Show($"Error initializing webcam: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        // WEBCAM DISCOVERY AND MANAGEMENT
        private List<string> GetWebcamNames()
        {
            var webcams = new List<string>();
            try
            {
                var videoDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);
                foreach (var device in videoDevices)
                {
                    webcams.Add(device.Name);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting webcam names with DirectShowLib: {ex.Message}");
                return DiscoverWebcamsFallback();
            }
            return webcams;
        }

        private List<string> DiscoverWebcamsFallback()
        {
            var webcams = new List<string>();
            
            for (int i = 0; i < 6; i++)
            {
                VideoCapture testCapture = null;
                try
                {
                    testCapture = new VideoCapture(i);
                    System.Threading.Thread.Sleep(100);
                    
                    if (testCapture.IsOpened())
                    {
                        using (var testFrame = new Mat())
                        {
                            if (testCapture.Read(testFrame) && !testFrame.Empty())
                            {
                                webcams.Add($"Webcam {i}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Webcam {i} test failed: {ex.Message}");
                    continue;
                }
                finally
                {
                    try
                    {
                        testCapture?.Release();
                        testCapture?.Dispose();
                    }
                    catch { }
                }
            }
            
            return webcams;
        }

        // DIRECTSHOW RESOLUTION DETECTION
        private List<VideoFormat> GetSupportedVideoFormats(int deviceIndex)
        {
            var formats = new List<VideoFormat>();
            
            try
            {
                var videoDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);
                
                if (deviceIndex >= videoDevices.Length)
                    return formats;
                
                var device = videoDevices[deviceIndex];
                
                IGraphBuilder graphBuilder = null;
                ICaptureGraphBuilder2 captureGraphBuilder = null;
                IBaseFilter sourceFilter = null;
                
                try
                {
                    graphBuilder = (IGraphBuilder)new FilterGraph();
                    captureGraphBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();
                    
                    int hr = captureGraphBuilder.SetFiltergraph(graphBuilder);
                    DsError.ThrowExceptionForHR(hr);
                    
                    // Create the source filter
                    Guid iid = typeof(IBaseFilter).GUID;
                    object filterObject;
                    device.Mon.BindToObject(null, null, ref iid, out filterObject);
                    sourceFilter = (IBaseFilter)filterObject;
                    
                    hr = graphBuilder.AddFilter(sourceFilter, "Video Source");
                    DsError.ThrowExceptionForHR(hr);
                    
                    IPin outputPin = null;
                    hr = captureGraphBuilder.FindPin(sourceFilter, PinDirection.Output, PinCategory.Capture, MediaType.Video, false, 0, out outputPin);
                    
                    if (hr != 0 || outputPin == null)
                    {
                        outputPin = DsFindPin.ByDirection(sourceFilter, PinDirection.Output, 0);
                    }
                    
                    if (outputPin != null)
                    {
                        IAMStreamConfig streamConfig = null;
                        try
                        {
                            streamConfig = outputPin as IAMStreamConfig;
                            if (streamConfig == null)
                            {
                                IntPtr streamConfigPtr;
                                Guid streamConfigGuid = typeof(IAMStreamConfig).GUID;
                                hr = Marshal.QueryInterface(Marshal.GetIUnknownForObject(outputPin), ref streamConfigGuid, out streamConfigPtr);
                                if (hr == 0 && streamConfigPtr != IntPtr.Zero)
                                {
                                    streamConfig = (IAMStreamConfig)Marshal.GetObjectForIUnknown(streamConfigPtr);
                                    Marshal.Release(streamConfigPtr);
                                }
                            }
                            
                            if (streamConfig != null)
                            {
                                int count, size;
                                hr = streamConfig.GetNumberOfCapabilities(out count, out size);
                                
                                if (hr == 0)
                                {
                                    for (int i = 0; i < count; i++)
                                    {
                                        try
                                        {
                                            AMMediaType mediaType = null;
                                            IntPtr capsPtr = Marshal.AllocHGlobal(size);
                                            
                                            try
                                            {
                                                hr = streamConfig.GetStreamCaps(i, out mediaType, capsPtr);
                                                
                                                if (hr == 0 && mediaType != null && mediaType.majorType == MediaType.Video)
                                                {
                                                    if (mediaType.formatType == DirectShowLib.FormatType.VideoInfo)
                                                    {
                                                        var videoInfo = (VideoInfoHeader)Marshal.PtrToStructure(mediaType.formatPtr, typeof(VideoInfoHeader));
                                                        
                                                        var format = new VideoFormat
                                                        {
                                                            Width = videoInfo.BmiHeader.Width,
                                                            Height = Math.Abs(videoInfo.BmiHeader.Height),
                                                            PixelFormat = GetPixelFormatName(mediaType.subType),
                                                            FrameRate = videoInfo.AvgTimePerFrame > 0 ? 10000000.0 / videoInfo.AvgTimePerFrame : 0
                                                        };
                                                        
                                                        if (format.Width > 0 && format.Height > 0)
                                                        {
                                                            formats.Add(format);
                                                        }
                                                    }
                                                    else if (mediaType.formatType == DirectShowLib.FormatType.VideoInfo2)
                                                    {
                                                        var videoInfo2 = (VideoInfoHeader2)Marshal.PtrToStructure(mediaType.formatPtr, typeof(VideoInfoHeader2));
                                                        
                                                        var format = new VideoFormat
                                                        {
                                                            Width = videoInfo2.BmiHeader.Width,
                                                            Height = Math.Abs(videoInfo2.BmiHeader.Height),
                                                            PixelFormat = GetPixelFormatName(mediaType.subType),
                                                            FrameRate = videoInfo2.AvgTimePerFrame > 0 ? 10000000.0 / videoInfo2.AvgTimePerFrame : 0
                                                        };
                                                        
                                                        if (format.Width > 0 && format.Height > 0)
                                                        {
                                                            formats.Add(format);
                                                        }
                                                    }
                                                }
                                                
                                                if (mediaType != null)
                                                {
                                                    DsUtils.FreeAMMediaType(mediaType);
                                                }
                                            }
                                            finally
                                            {
                                                Marshal.FreeHGlobal(capsPtr);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Error processing capability {i}: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (streamConfig != null)
                            {
                                Marshal.ReleaseComObject(streamConfig);
                            }
                        }
                        
                        Marshal.ReleaseComObject(outputPin);
                    }
                }
                finally
                {
                    if (sourceFilter != null) Marshal.ReleaseComObject(sourceFilter);
                    if (captureGraphBuilder != null) Marshal.ReleaseComObject(captureGraphBuilder);
                    if (graphBuilder != null) Marshal.ReleaseComObject(graphBuilder);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DirectShow enumeration error: {ex.Message}");
            }
            
            return formats;
        }

        private string GetPixelFormatName(Guid subType)
        {
            if (subType == MediaSubType.RGB24) return "RGB24";
            if (subType == MediaSubType.RGB32) return "RGB32";
            if (subType == MediaSubType.ARGB32) return "ARGB32";
            if (subType == MediaSubType.YUY2) return "YUY2";
            if (subType == MediaSubType.UYVY) return "UYVY";
            if (subType == MediaSubType.MJPG) return "MJPEG";
            if (subType == MediaSubType.I420) return "I420";
            if (subType == MediaSubType.YV12) return "YV12";
            if (subType == MediaSubType.NV12) return "NV12";
            
            try
            {
                byte[] bytes = subType.ToByteArray();
                return System.Text.Encoding.ASCII.GetString(bytes, 0, 4);
            }
            catch
            {
                return "Unknown";
            }
        }

        private List<Size> GetSupportedResolutions(VideoCapture capture)
        {
            var resolutions = new HashSet<Size>();
            
            if (capture == null || currentDeviceIndex < 0 || currentDeviceIndex >= availableWebcams.Count)
                return new List<Size>();

            try
            {
                var formats = GetSupportedVideoFormats(currentDeviceIndex);
                
                if (formats.Count > 0)
                {
                    foreach (var format in formats)
                    {
                        if (format.Width > 0 && format.Height > 0)
                        {
                            resolutions.Add(new Size(format.Width, format.Height));
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"Found {formats.Count} video formats, {resolutions.Count} unique resolutions via DirectShow");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("DirectShow format detection failed, using fallback method");
                    return GetSupportedResolutionsFallback(capture);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting DirectShow formats: {ex.Message}");
                return GetSupportedResolutionsFallback(capture);
            }

            var sortedResolutions = resolutions.ToList();
            sortedResolutions.Sort((a, b) => 
            {
                int areaA = a.Width * a.Height;
                int areaB = b.Width * b.Height;
                return areaB.CompareTo(areaA);
            });
            
            return sortedResolutions;
        }

        private List<Size> GetSupportedResolutionsFallback(VideoCapture capture)
        {
            var resolutions = new HashSet<Size>();
            
            if (capture == null || !capture.IsOpened()) 
                return new List<Size>();

            var originalWidth = (int)capture.Get(VideoCaptureProperties.FrameWidth);
            var originalHeight = (int)capture.Get(VideoCaptureProperties.FrameHeight);
            resolutions.Add(new Size(originalWidth, originalHeight));
            
            var commonResolutions = new List<Size>
            {
                new Size(1920, 1080),
                new Size(1280, 720),
                new Size(800, 600),
                new Size(640, 480),
                new Size(320, 240),
            };
            
            foreach (var testRes in commonResolutions)
            {
                try
                {
                    capture.Set(VideoCaptureProperties.FrameWidth, testRes.Width);
                    capture.Set(VideoCaptureProperties.FrameHeight, testRes.Height);
                    System.Threading.Thread.Sleep(50);
                    
                    var actualWidth = (int)capture.Get(VideoCaptureProperties.FrameWidth);
                    var actualHeight = (int)capture.Get(VideoCaptureProperties.FrameHeight);
                    
                    if (Math.Abs(actualWidth - testRes.Width) <= 10 && 
                        Math.Abs(actualHeight - testRes.Height) <= 10)
                    {
                        resolutions.Add(new Size(actualWidth, actualHeight));
                    }
                }
                catch
                {
                    continue;
                }
            }
            
            try
            {
                capture.Set(VideoCaptureProperties.FrameWidth, originalWidth);
                capture.Set(VideoCaptureProperties.FrameHeight, originalHeight);
            }
            catch { }
            
            return resolutions.ToList();
        }

        // WEBCAM CONTROL
        private bool StartWebcam(int deviceIndex)
        {
            if (isDisposing) return false;
            
            try
            {
                lock (lockObject)
                {
                    StopWebcam();
                    
                    if (deviceIndex >= 0 && deviceIndex < availableWebcams.Count)
                    {
                        currentDeviceIndex = deviceIndex;
                        
                        videoCapture = new VideoCapture(deviceIndex);
                        
                        if (!videoCapture.IsOpened())
                        {
                            throw new Exception("Failed to open webcam - it may be in use by another application");
                        }
                        
                        try
                        {
                            videoCapture.Set(VideoCaptureProperties.FrameWidth, 640);
                            videoCapture.Set(VideoCaptureProperties.FrameHeight, 480);
                            videoCapture.Set(VideoCaptureProperties.Fps, 30);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Could not set initial video properties: {ex.Message}");
                        }

                        using (var testFrame = new Mat())
                        {
                            if (!videoCapture.Read(testFrame) || testFrame.Empty())
                            {
                                throw new Exception("Webcam opened but cannot read frames");
                            }
                        }

                        supportedResolutions.Clear();
                        
                        // Get supported resolutions and automatically select the smallest one
                        try
                        {
                            ShowSwitchingScreen("DETECTING RESOLUTIONS...");
                            supportedResolutions = GetSupportedResolutions(videoCapture);
                            
                            if (supportedResolutions.Count > 0)
                            {
                                var smallestResolution = supportedResolutions
                                    .OrderBy(r => r.Width * r.Height)
                                    .First();
                                
                                videoCapture.Set(VideoCaptureProperties.FrameWidth, smallestResolution.Width);
                                videoCapture.Set(VideoCaptureProperties.FrameHeight, smallestResolution.Height);
                                
                                System.Threading.Thread.Sleep(100);
                                
                                var actualWidth = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
                                var actualHeight = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
                                currentResolution = new Size(actualWidth, actualHeight);
                                
                                System.Diagnostics.Debug.WriteLine($"Auto-selected smallest resolution: {actualWidth}x{actualHeight}");
                            }
                            else
                            {
                                var width = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
                                var height = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
                                currentResolution = new Size(width, height);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error setting smallest resolution: {ex.Message}");
                            var width = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
                            var height = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
                            currentResolution = new Size(width, height);
                        }
                        
                        // If we get here, webcam started successfully - remove from failed list
                        failedWebcamIndices.Remove(deviceIndex);
                        
                        UpdateWindowTitle(availableWebcams[deviceIndex]);
                        UpdateTrayMenu();
                        ClearDisabledScreen();
                        frameTimer.Start();
                        
                        System.Diagnostics.Debug.WriteLine($"Started webcam {deviceIndex}: {availableWebcams[deviceIndex]} at {currentResolution.Width}x{currentResolution.Height}");
                        
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting webcam: {ex.Message}");
                
                // Mark this webcam as failed
                failedWebcamIndices.Add(deviceIndex);
                
                try
                {
                    videoCapture?.Release();
                    videoCapture?.Dispose();
                    videoCapture = null;
                }
                catch { }
                
                throw new Exception($"Failed to start webcam '{availableWebcams[deviceIndex]}': {ex.Message}");
            }
            
            return false;
        }

        private void StopWebcam()
        {
            if (isDisposing) return;
            
            try
            {
                lock (lockObject)
                {
                    frameTimer?.Stop();
                    
                    if (videoCapture != null)
                    {
                        try
                        {
                            if (videoCapture.IsOpened())
                            {
                                videoCapture.Release();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error releasing video capture: {ex.Message}");
                        }
                        
                        try
                        {
                            videoCapture.Dispose();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error disposing video capture: {ex.Message}");
                        }
                        
                        videoCapture = null;
                    }
                    
                    if (!isDisposing && videoPanel != null && !videoPanel.IsDisposed)
                    {
                        if (videoPanel.InvokeRequired)
                        {
                            try
                            {
                                videoPanel.Invoke(new Action(() => ClearVideoPanel()));
                            }
                            catch (ObjectDisposedException)
                            {
                                // Form is being disposed, ignore
                            }
                        }
                        else
                        {
                            ClearVideoPanel();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping webcam: {ex.Message}");
            }

            if (!isDisposing)
            {
                currentWebcamName = "";
                UpdateWindowTitle("");
            }
        }

        // VIDEO DISPLAY
        private void FrameTimer_Tick(object sender, EventArgs e)
        {
            if (isDisposing || videoCapture == null || !isWebcamEnabled) 
                return;
            
            try
            {
                lock (lockObject)
                {
                    if (videoCapture != null && videoCapture.IsOpened())
                    {
                        using (var frame = new Mat())
                        {
                            if (videoCapture.Read(frame) && !frame.Empty())
                            {
                                DisplayFrame(frame);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Frame capture error: {ex.Message}");
                
                if (ex.Message.Contains("disposed") || ex.Message.Contains("released"))
                {
                    frameTimer?.Stop();
                }
            }
        }

        private void DisplayFrame(Mat frame)
        {
            if (isDisposing || frame == null || frame.Empty()) 
                return;
            
            try
            {
                if (videoPanel != null && !videoPanel.IsDisposed)
                {
                    if (videoPanel.InvokeRequired)
                    {
                        try
                        {
                            videoPanel.BeginInvoke(new Action(() => DisplayFrame(frame)));
                        }
                        catch (ObjectDisposedException)
                        {
                            // Form is being disposed, ignore
                        }
                        return;
                    }
                    
                    using (var bitmap = BitmapConverter.ToBitmap(frame))
                    {
                        var oldImage = videoPanel.BackgroundImage;
                        videoPanel.BackgroundImage = new Bitmap(bitmap);
                        videoPanel.BackgroundImageLayout = ImageLayout.Zoom;
                        oldImage?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Display frame error: {ex.Message}");
            }
        }

        private void ClearVideoPanel()
        {
            try
            {
                if (videoPanel != null && !videoPanel.IsDisposed)
                {
                    var oldImage = videoPanel.BackgroundImage;
                    videoPanel.BackgroundImage = null;
                    oldImage?.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing video panel: {ex.Message}");
            }
        }

        // AUTO-RESIZE FUNCTIONALITY
        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            if (this.WindowState == FormWindowState.Minimized)
            {
                if (isWebcamEnabled)
                {
                    wasDisabledByClose = true;
                    DisableWebcam();
                }

                this.Hide();
                trayIcon.Visible = true;
                trayIcon.ShowBalloonTip(2000, "Webcam Viewer", "Application minimized to tray", ToolTipIcon.Info);
            }
            else
            {
                if (!isBorderless && videoPanel != null)
                {
                    videoPanel.Size = this.ClientSize;
                }
                
                // Trigger auto-resize timer if conditions are met
                if (ShouldTriggerAutoResize())
                {
                    // Store location at the start of resize sequence
                    if (!isInResizeSequence)
                    {
                        preResizeLocation = this.Location;
                        isInResizeSequence = true;
                        System.Diagnostics.Debug.WriteLine($"Stored pre-resize location: {preResizeLocation}");
                    }
                    
                    // Ensure autoResizeTimer is initialized
                    if (autoResizeTimer == null)
                    {
                        autoResizeTimer = new Timer();
                        autoResizeTimer.Interval = 5000;
                        autoResizeTimer.Tick += AutoResizeTimer_Tick;
                    }
                    
                    // Reset timer on each resize event
                    autoResizeTimer.Stop();
                    autoResizeTimer.Start();
                    System.Diagnostics.Debug.WriteLine("Auto-resize timer restarted (5 seconds from resize completion)");
                }
            }
        }

        private bool ShouldTriggerAutoResize()
        {
            return !isDisposing && 
                   !isAutoResizing && 
                   !isFullScreen && 
                   !isBorderless && 
                   isWebcamEnabled && 
                   currentResolution.Width > 0 && 
                   currentResolution.Height > 0;
        }

        private void AutoResizeTimer_Tick(object sender, EventArgs e)
        {
            autoResizeTimer.Stop();
            
            if (!ShouldTriggerAutoResize())
            {
                System.Diagnostics.Debug.WriteLine("Auto-resize cancelled - conditions not met");
                isInResizeSequence = false;
                return;
            }
            
            try
            {
                isAutoResizing = true;
                PerformAutoResize();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Auto-resize error: {ex.Message}");
            }
            finally
            {
                isAutoResizing = false;
                isInResizeSequence = false;
            }
        }

        private void PerformAutoResize()
        {
            if (currentResolution.Width <= 0 || currentResolution.Height <= 0)
            {
                System.Diagnostics.Debug.WriteLine("Invalid resolution for auto-resize");
                return;
            }

            var clientSize = this.ClientSize;
            double videoAspectRatio = (double)currentResolution.Width / currentResolution.Height;
            double clientAspectRatio = (double)clientSize.Width / clientSize.Height;
            
            System.Diagnostics.Debug.WriteLine($"Video aspect ratio: {videoAspectRatio:F3}, Client aspect ratio: {clientAspectRatio:F3}");
            
            if (Math.Abs(videoAspectRatio - clientAspectRatio) < 0.01)
            {
                System.Diagnostics.Debug.WriteLine("Aspect ratios already match, no resize needed");
                return;
            }

            Size newClientSize;
            
            if (videoAspectRatio > clientAspectRatio)
            {
                newClientSize = new Size(clientSize.Width, (int)(clientSize.Width / videoAspectRatio));
            }
            else
            {
                newClientSize = new Size((int)(clientSize.Height * videoAspectRatio), clientSize.Height);
            }
            
            var currentWindowSize = this.Size;
            var currentClientSize = this.ClientSize;
            var chromeWidth = currentWindowSize.Width - currentClientSize.Width;
            var chromeHeight = currentWindowSize.Height - currentClientSize.Height;
            
            var newWindowSize = new Size(
                newClientSize.Width + chromeWidth,
                newClientSize.Height + chromeHeight
            );
            
            var screen = Screen.FromControl(this);
            var workingArea = screen.WorkingArea;
            
            if (newWindowSize.Width > workingArea.Width || newWindowSize.Height > workingArea.Height)
            {
                double scaleX = (double)workingArea.Width / newWindowSize.Width;
                double scaleY = (double)workingArea.Height / newWindowSize.Height;
                double scale = Math.Min(scaleX, scaleY);
                
                newWindowSize = new Size(
                    (int)(newWindowSize.Width * scale),
                    (int)(newWindowSize.Height * scale)
                );
            }
            
            System.Drawing.Point newLocation;
            
            // Check if this is the first auto-resize or if preResizeLocation is invalid
            if (preResizeLocation.X <= 10 && preResizeLocation.Y <= 10)
            {
                // First time or invalid location - center the window
                newLocation = new System.Drawing.Point(
                    workingArea.X + (workingArea.Width - newWindowSize.Width) / 2,
                    workingArea.Y + (workingArea.Height - newWindowSize.Height) / 2
                );
                System.Diagnostics.Debug.WriteLine("First auto-resize - centering window");
            }
            else
            {
                // Try to preserve the top-left position from before resize
                newLocation = preResizeLocation;
                
                // Ensure the window stays within screen bounds
                if (newLocation.X + newWindowSize.Width > workingArea.Right)
                {
                    newLocation.X = workingArea.Right - newWindowSize.Width;
                }
                if (newLocation.Y + newWindowSize.Height > workingArea.Bottom)
                {
                    newLocation.Y = workingArea.Bottom - newWindowSize.Height;
                }
                if (newLocation.X < workingArea.Left)
                {
                    newLocation.X = workingArea.Left;
                }
                if (newLocation.Y < workingArea.Top)
                {
                    newLocation.Y = workingArea.Top;
                }
                System.Diagnostics.Debug.WriteLine($"Preserving top-left position: {newLocation}");
            }
            
            System.Diagnostics.Debug.WriteLine($"Auto-resizing window from {this.Size} to {newWindowSize}");
            System.Diagnostics.Debug.WriteLine($"Final window location: {newLocation}");
            
            this.SetBounds(newLocation.X, newLocation.Y, newWindowSize.Width, newWindowSize.Height);
        }

        // WINDOW TITLE AND STATUS
        private void ClockTimer_Tick(object sender, EventArgs e)
        {
            UpdateWindowTitle(currentWebcamName);
        }

        private void UpdateWindowTitle(string webcamName)
        {
            if (isDisposing) return;
            
            try
            {
                currentWebcamName = webcamName;
                string baseTitle = string.IsNullOrEmpty(webcamName) ? "Webcam Viewer" : $"Webcam Viewer - {webcamName}";
                string status = isWebcamEnabled ? string.Empty : " [DISABLED]";
                string clock = DateTime.Now.ToString("[HH:mm:ss] ");
                
                this.Text = $"{clock}{baseTitle}{status}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating window title: {ex.Message}");
            }
        }

        private void ShowDisabledScreen()
        {
            ShowStatusScreen("WEBCAM DISABLED", Color.DarkGray);
        }

        private void ShowSwitchingScreen(string message = "SWITCHING WEBCAM...")
        {
            ShowStatusScreen(message, Color.Navy);
        }

        private void ShowStatusScreen(string message, Color backgroundColor)
        {
            if (isDisposing) return;
            
            try
            {
                if (videoPanel.InvokeRequired)
                {
                    videoPanel.BeginInvoke(new Action(() => ShowStatusScreen(message, backgroundColor)));
                    return;
                }

                if (videoPanel != null && !videoPanel.IsDisposed)
                {
                    var oldImage = videoPanel.BackgroundImage;
                    
                    int width = Math.Max(videoPanel.Width, 1);
                    int height = Math.Max(videoPanel.Height, 1);
                    
                    Bitmap statusImage = new Bitmap(width, height);
                    using (Graphics g = Graphics.FromImage(statusImage))
                    {
                        using (SolidBrush brush = new SolidBrush(backgroundColor))
                        {
                            g.FillRectangle(brush, 0, 0, width, height);
                        }
                        
                        using (SolidBrush textBrush = new SolidBrush(Color.White))
                        using (Font font = new Font("Arial", 20, FontStyle.Bold))
                        {
                            SizeF textSize = g.MeasureString(message, font);
                            PointF textLocation = new PointF(
                                (width - textSize.Width) / 2,
                                (height - textSize.Height) / 2
                            );
                            g.DrawString(message, font, textBrush, textLocation);
                        }
                    }
                    
                    videoPanel.BackgroundImage = statusImage;
                    videoPanel.BackgroundImageLayout = ImageLayout.Center;
                    
                    oldImage?.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error showing status screen: {ex.Message}");
            }
        }

        private void ClearDisabledScreen()
        {
            if (isDisposing) return;
            
            try
            {
                if (videoPanel.InvokeRequired)
                {
                    videoPanel.BeginInvoke(new Action(ClearDisabledScreen));
                    return;
                }

                if (videoPanel != null && !videoPanel.IsDisposed)
                {
                    var oldImage = videoPanel.BackgroundImage;
                    videoPanel.BackgroundImage = null;
                    videoPanel.BackColor = Color.Black;
                    oldImage?.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing disabled screen: {ex.Message}");
            }
        }

        // WEBCAM STATE MANAGEMENT
        private void EnableWebcam()
        {
            if (isDisposing) return;
            
            if (!isWebcamEnabled)
            {
                isWebcamEnabled = true;
                UpdateWindowTitle(currentWebcamName);
                
                ClearDisabledScreen();
                
                if (availableWebcams != null && currentDeviceIndex >= 0 && currentDeviceIndex < availableWebcams.Count)
                {
                    try
                    {
                        StartWebcam(currentDeviceIndex);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error starting webcam: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        isWebcamEnabled = false;
                        ShowDisabledScreen();
                        UpdateWindowTitle(currentWebcamName);
                    }
                }
                
                UpdateTrayMenu();
            }
        }

        private void DisableWebcam()
        {
            if (isDisposing) return;
            
            if (isWebcamEnabled)
            {
                isWebcamEnabled = false;
                UpdateWindowTitle(currentWebcamName);
                ShowDisabledScreen();
                UpdateTrayMenu();
            }
        }

        // TRAY ICON AND MENU
        private void SetupTrayIcon()
        {
            UpdateTrayMenu();
        }

        private void UpdateTrayMenu()
        {
            if (isDisposing) return;
            
            try
            {
                trayContextMenu.Items.Clear();
                
                trayContextMenu.ImageScalingSize = new System.Drawing.Size(16, 16);
                
                Font baseFont = new Font("Segoe UI", 9F, FontStyle.Regular);
                Font boldFont = new Font("Segoe UI", 9F, FontStyle.Bold);
                
                trayContextMenu.Font = baseFont;

                // Webcam selection submenu
                ToolStripMenuItem webcamMenu = new ToolStripMenuItem("Select Webcam");
                webcamMenu.Font = baseFont;
                
                if (availableWebcams != null && availableWebcams.Count > 0)
                {
                    for (int i = 0; i < availableWebcams.Count; i++)
                    {
                        ToolStripMenuItem webcamItem = new ToolStripMenuItem();
                        
                        bool isCurrentWebcam = (i == currentDeviceIndex && videoCapture != null && !isDisposing);
                        bool isFailed = failedWebcamIndices.Contains(i);
                        
                        string webcamName = availableWebcams[i];
                        
                        if (isFailed)
                        {
                            webcamName += " [ERROR]";
                        }
                        
                        if (isCurrentWebcam)
                        {
                            webcamItem.Text = "✓ " + webcamName;
                            webcamItem.Font = boldFont;
                        }
                        else
                        {
                            webcamItem.Text = "    " + webcamName;
                            webcamItem.Font = baseFont;
                        }
                        
                        // Disable failed webcams so they can't be selected
                        if (isFailed)
                        {
                            webcamItem.Enabled = false;
                            webcamItem.ForeColor = Color.Gray;
                        }
                        
                        int index = i;
                        webcamItem.Click += (s, e) => WebcamMenuItem_Click(index);
                        webcamMenu.DropDownItems.Add(webcamItem);
                    }
                }
                else
                {
                    ToolStripMenuItem noWebcamItem = new ToolStripMenuItem("No webcams found");
                    noWebcamItem.Enabled = false;
                    noWebcamItem.Font = baseFont;
                    webcamMenu.DropDownItems.Add(noWebcamItem);
                }
                
                trayContextMenu.Items.Add(webcamMenu);

                // Resolution submenu
                ToolStripMenuItem resolutionMenu = new ToolStripMenuItem("Resolution");
                resolutionMenu.Font = baseFont;
                if (videoCapture == null || !isWebcamEnabled)
                {
                    resolutionMenu.Enabled = false;
                }
                else
                {
                    AddResolutionsToMenu(resolutionMenu);
                }
                trayContextMenu.Items.Add(resolutionMenu);
                
                // Refresh webcams
                ToolStripMenuItem refreshItem = new ToolStripMenuItem("Refresh Webcams");
                refreshItem.Font = baseFont;
                refreshItem.Click += RefreshMenuItem_Click;
                trayContextMenu.Items.Add(refreshItem);

                trayContextMenu.Items.Add(new ToolStripSeparator());

                // Enable/Disable webcam
                ToolStripMenuItem enableDisableItem = new ToolStripMenuItem(isWebcamEnabled ? "Disable Webcam" : "Enable Webcam");
                enableDisableItem.Font = baseFont;
                enableDisableItem.Click += EnableDisableMenuItem_Click;
                trayContextMenu.Items.Add(enableDisableItem);

                // Full Screen
                ToolStripMenuItem fullScreenItem = new ToolStripMenuItem();
                if (isFullScreen)
                {
                    fullScreenItem.Text = "✓ Full Screen";
                    fullScreenItem.Font = boldFont;
                }
                else
                {
                    fullScreenItem.Text = "    Full Screen";
                    fullScreenItem.Font = baseFont;
                }
                fullScreenItem.Click += FullScreenMenuItem_Click;
                trayContextMenu.Items.Add(fullScreenItem);
                
                // Always on top
                ToolStripMenuItem alwaysOnTopItem = new ToolStripMenuItem();
                if (this.TopMost)
                {
                    alwaysOnTopItem.Text = "✓ Always on Top";
                    alwaysOnTopItem.Font = boldFont;
                }
                else
                {
                    alwaysOnTopItem.Text = "    Always on Top";
                    alwaysOnTopItem.Font = baseFont;
                }
                alwaysOnTopItem.Click += AlwaysOnTopMenuItem_Click;
                trayContextMenu.Items.Add(alwaysOnTopItem);

                // Minimize to tray
                ToolStripMenuItem minimizeToTrayItem = new ToolStripMenuItem();
                if (minimizeToTrayEnabled)
                {
                    minimizeToTrayItem.Text = "✓ Minimize to Tray";
                    minimizeToTrayItem.Font = boldFont;
                }
                else
                {
                    minimizeToTrayItem.Text = "    Minimize to tray";
                    minimizeToTrayItem.Font = baseFont;
                }
                minimizeToTrayItem.Click += MinimizeToTrayMenuItem_Click;
                trayContextMenu.Items.Add(minimizeToTrayItem);

                trayContextMenu.Items.Add(new ToolStripSeparator());

                // Diagnostics
                ToolStripMenuItem diagnosticsItem = new ToolStripMenuItem("Webcam Diagnostics");
                diagnosticsItem.Font = baseFont;
                diagnosticsItem.Click += DiagnosticsMenuItem_Click;
                trayContextMenu.Items.Add(diagnosticsItem);

                trayContextMenu.Items.Add(new ToolStripSeparator());

                // Exit
                ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
                exitItem.Font = baseFont;
                exitItem.Click += ExitMenuItem_Click;
                trayContextMenu.Items.Add(exitItem);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating tray menu: {ex.Message}");
            }
        }

        private void AddResolutionsToMenu(ToolStripMenuItem parentMenu)
        {
            parentMenu.DropDownItems.Clear();
            
            Font baseFont = new Font("Segoe UI", 9F, FontStyle.Regular);
            Font boldFont = new Font("Segoe UI", 9F, FontStyle.Bold);
            
            if (videoCapture == null || !videoCapture.IsOpened() || !isWebcamEnabled)
            {
                ToolStripMenuItem disabledItem = new ToolStripMenuItem("Webcam not active");
                disabledItem.Enabled = false;
                disabledItem.Font = baseFont;
                parentMenu.DropDownItems.Add(disabledItem);
                return;
            }
            
            if (supportedResolutions.Count == 0)
            {
                try
                {
                    ShowSwitchingScreen("DETECTING RESOLUTIONS...");
                    supportedResolutions = GetSupportedResolutions(videoCapture);
                    ClearDisabledScreen();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting resolutions: {ex.Message}");
                    ToolStripMenuItem errorItem = new ToolStripMenuItem("Error detecting resolutions");
                    errorItem.Enabled = false;
                    errorItem.Font = baseFont;
                    parentMenu.DropDownItems.Add(errorItem);
                    return;
                }
            }
            
            if (supportedResolutions.Count == 0)
            {
                ToolStripMenuItem noResItem = new ToolStripMenuItem("No resolutions detected");
                noResItem.Enabled = false;
                noResItem.Font = baseFont;
                parentMenu.DropDownItems.Add(noResItem);
                return;
            }
            
            var currentWidth = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
            var currentHeight = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
            currentResolution = new Size(currentWidth, currentHeight);
            
            foreach (var res in supportedResolutions)
            {
                ToolStripMenuItem resItem = new ToolStripMenuItem();
                string resText = $"{res.Width}x{res.Height}";
                
                bool isCurrent = (Math.Abs(res.Width - currentWidth) <= 5 && 
                                 Math.Abs(res.Height - currentHeight) <= 5);
                
                if (isCurrent)
                {
                    resItem.Text = "✓ " + resText;
                    resItem.Font = boldFont;
                }
                else
                {
                    resItem.Text = "    " + resText;
                    resItem.Font = baseFont;
                }
                
                resItem.Tag = res;
                resItem.Click += ResolutionMenuItem_Click;
                parentMenu.DropDownItems.Add(resItem);
            }
            
            parentMenu.DropDownItems.Add(new ToolStripSeparator());
            ToolStripMenuItem refreshResItem = new ToolStripMenuItem("Refresh Resolutions");
            refreshResItem.Font = baseFont;
            refreshResItem.Click += RefreshResolutions_Click;
            parentMenu.DropDownItems.Add(refreshResItem);
        }

        // MENU EVENT HANDLERS
        private async void WebcamMenuItem_Click(int deviceIndex)
        {
            if (isDisposing) return;
            
            if (deviceIndex == currentDeviceIndex && videoCapture != null && isWebcamEnabled)
            {
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;
                ShowSwitchingScreen();
                
                if (isWebcamEnabled)
                {
                    isWebcamEnabled = false;
                    UpdateWindowTitle(currentWebcamName);
                }
                
                StopWebcam();
                
                await Task.Delay(300);
                
                currentDeviceIndex = deviceIndex;
                
                EnableWebcam();
                
                UpdateTrayMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error switching webcam: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DisableWebcam();
                UpdateTrayMenu();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private async void RefreshMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            try
            {
                this.Cursor = Cursors.WaitCursor;
                ShowSwitchingScreen("REFRESHING...");
                
                bool wasEnabled = isWebcamEnabled;
                if (isWebcamEnabled)
                {
                    DisableWebcam();
                }
                StopWebcam();
                
                // Clear failed webcam list to allow retry
                failedWebcamIndices.Clear();
                System.Diagnostics.Debug.WriteLine("Cleared failed webcam list during refresh");
                
                await Task.Run(() => {
                    availableWebcams = GetWebcamNames();
                });
                
                if (currentDeviceIndex >= availableWebcams.Count)
                {
                    currentDeviceIndex = 0;
                }
                
                if (wasEnabled && availableWebcams.Count > 0)
                {
                    EnableWebcam();
                }
                
                UpdateTrayMenu();
                
                MessageBox.Show($"Found {availableWebcams.Count} webcam(s)\n\nFailed webcams have been re-enabled for retry.", 
                               "Webcam Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing webcams: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ResolutionMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing || !(sender is ToolStripMenuItem menuItem)) return;
            
            if (menuItem.Tag is Size selectedRes)
            {
                SetWebcamResolution(selectedRes.Width, selectedRes.Height);
            }
        }

        private void RefreshResolutions_Click(object sender, EventArgs e)
        {
            if (isDisposing || videoCapture == null || !isWebcamEnabled) return;
            
            try
            {
                this.Cursor = Cursors.WaitCursor;
                ShowSwitchingScreen("REFRESHING RESOLUTIONS...");
                
                supportedResolutions.Clear();
                supportedResolutions = GetSupportedResolutions(videoCapture);
                
                UpdateTrayMenu();
                
                MessageBox.Show($"Found {supportedResolutions.Count} supported resolution(s)", 
                               "Resolution Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing resolutions: {ex.Message}", 
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                ClearDisabledScreen();
            }
        }

        private void SetWebcamResolution(int width, int height)
        {
            if (isDisposing || videoCapture == null || !isWebcamEnabled) return;

            var currentWidth = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
            var currentHeight = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
            
            if (Math.Abs(currentWidth - width) <= 5 && Math.Abs(currentHeight - height) <= 5)
            {
                System.Diagnostics.Debug.WriteLine($"Resolution {width}x{height} is already active");
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;
                ShowSwitchingScreen($"CHANGING TO {width}x{height}...");
                
                lock (lockObject)
                {
                    if (videoCapture == null || !videoCapture.IsOpened())
                    {
                        throw new Exception("Webcam not available for resolution change");
                    }
                    
                    frameTimer?.Stop();
                    
                    videoCapture.Set(VideoCaptureProperties.FrameWidth, width);
                    videoCapture.Set(VideoCaptureProperties.FrameHeight, height);
                    
                    System.Threading.Thread.Sleep(200);
                    
                    var actualWidth = (int)videoCapture.Get(VideoCaptureProperties.FrameWidth);
                    var actualHeight = (int)videoCapture.Get(VideoCaptureProperties.FrameHeight);
                    
                    using (var testFrame = new Mat())
                    {
                        for (int attempts = 0; attempts < 3; attempts++)
                        {
                            if (videoCapture.Read(testFrame) && !testFrame.Empty())
                            {
                                break;
                            }
                            System.Threading.Thread.Sleep(100);
                        }
                        
                        if (testFrame.Empty())
                        {
                            throw new Exception($"Cannot capture frames at resolution {actualWidth}x{actualHeight}");
                        }
                    }
                    
                    currentResolution = new Size(actualWidth, actualHeight);
                    frameTimer?.Start();
                    
                    System.Diagnostics.Debug.WriteLine($"Resolution changed successfully to {actualWidth}x{actualHeight}");
                }
                
                UpdateTrayMenu();
                ClearDisabledScreen();
                
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Resolution change failed: {ex.Message}");
                
                try
                {
                    if (videoCapture != null && videoCapture.IsOpened() && isWebcamEnabled)
                    {
                        frameTimer?.Start();
                    }
                }
                catch { }
                
                MessageBox.Show($"Failed to change resolution to {width}x{height}:\n{ex.Message}\n\nThe webcam may not support this resolution.", 
                               "Resolution Change Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                UpdateTrayMenu();
                ClearDisabledScreen();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void EnableDisableMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            if (isWebcamEnabled)
            {
                DisableWebcam();
            }
            else
            {
                EnableWebcam();
            }
        }

        private void FullScreenMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            ToggleFullScreen();
        }

        private void AlwaysOnTopMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            this.TopMost = !this.TopMost;
            UpdateTrayMenu();
        }

        private void MinimizeToTrayMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            minimizeToTrayEnabled = !minimizeToTrayEnabled;
            UpdateTrayMenu();
        }

        private void DiagnosticsMenuItem_Click(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            try
            {
                string diagnostics = "=== WEBCAM DIAGNOSTICS (OpenCvSharp + DirectShow) ===\r\n\r\n";
                
                diagnostics += $"Available webcams: {availableWebcams?.Count ?? 0}\r\n";
                if (availableWebcams != null)
                {
                    for (int i = 0; i < availableWebcams.Count; i++)
                    {
                        string status = failedWebcamIndices.Contains(i) ? " [FAILED]" : "";
                        diagnostics += $"  {i}: {availableWebcams[i]}{status}\r\n";
                    }
                }
                
                diagnostics += $"\r\nFailed webcams: {failedWebcamIndices.Count}\r\n";
                if (failedWebcamIndices.Count > 0)
                {
                    diagnostics += $"Failed indices: [{string.Join(", ", failedWebcamIndices)}]\r\n";
                }
                
                diagnostics += $"\r\nCurrent webcam index: {currentDeviceIndex}\r\n";
                diagnostics += $"Webcam enabled: {isWebcamEnabled}\r\n";
                diagnostics += $"Current resolution: {currentResolution.Width}x{currentResolution.Height}\r\n";
                diagnostics += $"Supported resolutions: {supportedResolutions.Count}\r\n";
                
                lock (lockObject)
                {
                    if (videoCapture != null)
                    {
                        diagnostics += $"VideoCapture opened: {videoCapture.IsOpened()}\r\n";
                        
                        if (videoCapture.IsOpened())
                        {
                            try
                            {
                                double width = videoCapture.Get(VideoCaptureProperties.FrameWidth);
                                double height = videoCapture.Get(VideoCaptureProperties.FrameHeight);
                                double fps = videoCapture.Get(VideoCaptureProperties.Fps);
                                
                                diagnostics += $"Actual resolution: {width}x{height}\r\n";
                                diagnostics += $"Actual FPS: {fps}\r\n";
                            }
                            catch (Exception ex)
                            {
                                diagnostics += $"Error getting webcam properties: {ex.Message}\r\n";
                            }
                        }
                    }
                    else
                    {
                        diagnostics += "VideoCapture: null\r\n";
                    }
                }
                
                diagnostics += $"\r\nFrame timer running: {frameTimer?.Enabled ?? false}\r\n";
                diagnostics += $"Frame timer interval: {frameTimer?.Interval ?? 0}ms\r\n";
                diagnostics += $"Auto-resize enabled: {!isAutoResizing}\r\n";
                diagnostics += $"Auto-resize timer running: {autoResizeTimer?.Enabled ?? false}\r\n";
                diagnostics += $"In resize sequence: {isInResizeSequence}\r\n";
                diagnostics += $"Pre-resize location: {preResizeLocation}\r\n";
                
                diagnostics += "\r\n=== SYSTEM INFO ===\r\n";
                diagnostics += $"OpenCV Version: {Cv2.GetVersionString()}\r\n";
                diagnostics += $"OS: {Environment.OSVersion}\r\n";
                diagnostics += $".NET Version: {Environment.Version}\r\n";
                
                diagnostics += "\r\n=== TROUBLESHOOTING TIPS ===\r\n";
                diagnostics += "• Close other apps using webcam (Zoom, Teams, etc.)\r\n";
                diagnostics += "• Try 'Refresh Webcams' to detect new devices\r\n";
                diagnostics += "• Update webcam drivers in Device Manager\r\n";
                diagnostics += "• Try different USB port\r\n";
                diagnostics += "• Check Windows Camera app works\r\n";
                diagnostics += "• Restart application if issues persist\r\n";
                
                ShowDiagnosticsDialog(diagnostics);
                
                var result = MessageBox.Show("Show detailed video format information for current webcam?", 
                                           "Detailed Formats", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes && currentDeviceIndex >= 0)
                {
                    ShowDetailedVideoFormats(currentDeviceIndex);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting diagnostics: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowDetailedVideoFormats(int deviceIndex)
        {
            try
            {
                var formats = GetSupportedVideoFormats(deviceIndex);
                
                string info = $"=== DETAILED VIDEO FORMATS FOR WEBCAM {deviceIndex} ===\r\n";
                if (deviceIndex < availableWebcams.Count)
                {
                    info += $"Device: {availableWebcams[deviceIndex]}\r\n\r\n";
                }
                
                if (formats.Count > 0)
                {
                    var grouped = formats.GroupBy(f => new { f.Width, f.Height })
                                        .OrderByDescending(g => g.Key.Width * g.Key.Height);
                    
                    foreach (var group in grouped)
                    {
                        info += $"{group.Key.Width}x{group.Key.Height}:\r\n";
                        foreach (var format in group.OrderByDescending(f => f.FrameRate))
                        {
                            info += $"  • {format.PixelFormat} @ {format.FrameRate:F1} fps\r\n";
                        }
                        info += "\r\n";
                    }
                }
                else
                {
                    info += "No formats detected via DirectShow\r\n";
                    info += "This may indicate:\r\n";
                    info += "• Webcam driver issues\r\n";
                    info += "• Webcam is in use by another application\r\n";
                    info += "• DirectShow compatibility problems\r\n";
                }
                
                ShowDiagnosticsDialog(info);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting detailed formats: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowDiagnosticsDialog(string diagnostics)
        {
            try
            {
                Form diagForm = new Form()
                {
                    Text = "Webcam Diagnostics",
                    Size = new System.Drawing.Size(600, 700),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.Sizable,
                    MinimizeBox = false,
                    MaximizeBox = false
                };
                
                TextBox textBox = new TextBox()
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    ReadOnly = true,
                    Font = new Font("Consolas", 9),
                    Dock = DockStyle.Fill,
                    Text = diagnostics
                };
                
                Button closeBtn = new Button()
                {
                    Text = "Close",
                    DialogResult = DialogResult.OK,
                    Dock = DockStyle.Bottom,
                    Height = 35
                };
                
                diagForm.Controls.Add(textBox);
                diagForm.Controls.Add(closeBtn);
                diagForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing diagnostics dialog: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // WINDOW MODE MANAGEMENT
        private void VideoPanel_DoubleClick(object sender, EventArgs e)
        {
            if (isFullScreen) return;
            ToggleWindowMode();
        }

        private void ToggleWindowMode()
        {
            if (isDisposing) return;
            
            if (isBorderless)
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.WindowState = originalWindowState;
                this.Bounds = originalBounds;
                isBorderless = false;
            }
            else
            {
                originalBounds = this.Bounds;
                originalWindowState = this.WindowState;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Normal;
                isBorderless = true;
            }
        }

        private void ToggleFullScreen()
        {
            if (isFullScreen)
            {
                this.FormBorderStyle = originalBorderStyle;
                this.WindowState = originalWindowState;
                this.Bounds = originalBounds;
                isFullScreen = false;
            }
            else
            {
                originalBorderStyle = this.FormBorderStyle;
                originalWindowState = this.WindowState;
                originalBounds = this.Bounds;
                
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                isFullScreen = true;
            }
            UpdateTrayMenu();
        }

        // WINDOW EVENT HANDLERS
        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            if (isDisposing) return;
            
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();

            if (wasDisabledByClose)
            {
                wasDisabledByClose = false;
                EnableWebcam();
            }
        }

        private void MainForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (isDisposing) return;
            
            if (e.Button == MouseButtons.Right)
            {
                trayContextMenu.Show(this, e.Location);
            }
        }

        private void VideoPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (isDisposing) return;
            
            if (e.Button == MouseButtons.Right)
            {
                trayContextMenu.Show(videoPanel, e.Location);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (minimizeToTrayEnabled)
                {
                    e.Cancel = true;

                    if (isWebcamEnabled)
                    {
                        wasDisabledByClose = true;
                        DisableWebcam();
                    }

                    this.Hide();
                    trayIcon.ShowBalloonTip(2000, "Webcam Viewer", "Application minimized to tray", ToolTipIcon.Info);
                }
                else
                {
                    DialogResult result = MessageBox.Show(
                        "Are you sure you want to exit the application?", 
                        "Confirm Exit", 
                        MessageBoxButtons.YesNo, 
                        MessageBoxIcon.Question
                    );
                    
                    if (result == DialogResult.Yes)
                    {
                        e.Cancel = false;
                    }
                    else
                    {
                        e.Cancel = true;
                    }
                }
            }
            else
            {
                isDisposing = true;
                StopWebcam();
                trayIcon.Visible = false;
            }
        }

        protected override void Dispose(bool disposing)
        {
            isDisposing = true;
            
            if (disposing)
            {
                try
                {
                    frameTimer?.Stop();
                    frameTimer?.Dispose();
                    
                    clockTimer?.Stop();
                    clockTimer?.Dispose();
                    
                    autoResizeTimer?.Stop();
                    autoResizeTimer?.Dispose();
                    
                    StopWebcam();
                    
                    if (videoPanel?.BackgroundImage != null)
                    {
                        videoPanel.BackgroundImage.Dispose();
                        videoPanel.BackgroundImage = null;
                    }
                    
                    if (components != null)
                    {
                        components.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error during dispose: {ex.Message}");
                }
            }
            
            base.Dispose(disposing);
        }
    }
}