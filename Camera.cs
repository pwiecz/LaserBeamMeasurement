using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System;
using System.Drawing;

namespace LaserBeamMeasurement
{
    public interface ICamera
    {
        void Start();
        void Stop();
        Mat GetGrayFrame();
        event EventHandler OnNewFrame;
    }

    class VideoCaptureCamera : ICamera
    {
        public VideoCaptureCamera()
        {
            _capture = new VideoCapture();
            _capture.SetCaptureProperty(CapProp.FrameWidth, 1920);
            _capture.SetCaptureProperty(CapProp.FrameHeight, 1080);
        }

        public void Start()
        {
            _capture.Start();
        }
        public void Stop()
        {
            _capture.Pause();
        }
        public Mat GetGrayFrame()
        {
            Mat frame = new Mat();
            _capture.Retrieve(frame);
            Mat grayFrame = new Mat();
            CvInvoke.CvtColor(frame, grayFrame, ColorConversion.Bgr2Gray);
            return grayFrame;
        }
        event EventHandler ICamera.OnNewFrame
        {
            add
            {
                _capture.ImageGrabbed += value;
            }
            remove
            {
                _capture.ImageGrabbed -= value;
            }
        }

        VideoCapture _capture;
    }

    class UEyeCamera : ICamera
    {
        public bool Init()
        {
            uEye.Types.CameraInformation[] cameraList;
            uEye.Info.Camera.GetCameraList(out cameraList);
            foreach (uEye.Types.CameraInformation info in cameraList)
            {
                uEye.Camera camera = new uEye.Camera();
                if (camera.Init(info.DeviceID | (Int32)uEye.Defines.DeviceEnumeration.UseDeviceID) != uEye.Defines.Status.SUCCESS)
                {
                    return false;
                }
                if (MemoryHelper.AllocImageMems(camera, 3/*m_cnNumberOfSeqBuffers*/) != uEye.Defines.Status.SUCCESS)
                {
                    return false;
                }
                if (MemoryHelper.InitSequence(camera) != uEye.Defines.Status.SUCCESS)
                {
                    return false;
                }
                if (camera.AutoFeatures.Software.Shutter.SetEnable(true) != uEye.Defines.Status.SUCCESS)
                {
                    return false;
                }
                _camera = camera;
                return true;
            }
            return false;
        }
        public void Start()
        {
            _camera.Acquisition.Capture();
        }
        public void Stop()
        {
            _camera.Acquisition.Stop();
        }
        public Mat GetGrayFrame()
        {
            uEye.Defines.DisplayMode mode;
            _camera.Display.Mode.Get(out mode);

            // only display in dib mode
            if (mode != uEye.Defines.DisplayMode.DiB)
            {
                return null;
            }
            Int32 s32MemID;
            if (_camera.Memory.GetLast(out s32MemID) != uEye.Defines.Status.SUCCESS || s32MemID <= 0)
            {
                return null;
            }

            if (_camera.Memory.Lock(s32MemID) != uEye.Defines.Status.SUCCESS)
            {
                return null;
            }

            Bitmap bitmap;
            _camera.Memory.ToBitmap(s32MemID, out bitmap);

            _camera.Memory.Unlock(s32MemID);

            if (bitmap == null && bitmap.PixelFormat == System.Drawing.Imaging.PixelFormat.Format8bppIndexed)
            {
                return null;
            }


            Image<Bgr, byte> uEyeFrame = bitmap.ToImage<Bgr, byte>();

            Mat grayFrame = new Mat();
            CvInvoke.CvtColor(uEyeFrame, grayFrame, ColorConversion.Bgr2Gray);
            return grayFrame;

        }
        event EventHandler ICamera.OnNewFrame
        {
            add
            {
                _camera.EventFrame += value;
            }
            remove
            {
                _camera.EventFrame -= value;
            }
        }

        uEye.Camera _camera;
    }
}
