// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Drawing;
using Xunit;
using static Interop;
using static Interop.Ole32;
using static Interop.Oleaut32;

namespace System.Windows.Forms.Primitives.Tests.Interop.Oleaut32
{
    public class IDispatchTests
    {
        [WinFormsFact]
        public unsafe void IDispatch_GetIDsOfNames_Invoke_Success()
        {
            using var image = new Bitmap(16, 32);
            IPictureDisp picture = SubAxHost.GetIPictureDispFromPicture(image);
            IDispatch dispatch = (IDispatch)picture;

            Guid riid = Guid.Empty;
            var rgszNames = new string[] { "Width", "Other" };
            var rgDispId = new DispatchID[rgszNames.Length];
            fixed (DispatchID* pRgDispId = rgDispId)
            {
                HRESULT hr = dispatch.GetIDsOfNames(&riid, rgszNames, (uint)rgszNames.Length, 0, pRgDispId);
                Assert.Equal(HRESULT.S_OK, hr);
                Assert.Equal(new string[] { "Width", "Other" }, rgszNames);
                Assert.Equal(new DispatchID[] { (DispatchID)4, DispatchID.UNKNOWN }, rgDispId);
            }
        }

        [WinFormsFact]
        public unsafe void IDispatch_GetTypeInfo_Invoke_Success()
        {
            using var image = new Bitmap(16, 16);
            IPictureDisp picture = SubAxHost.GetIPictureDispFromPicture(image);
            IDispatch dispatch = (IDispatch)picture;

            ITypeInfo typeInfo;
            HRESULT hr = dispatch.GetTypeInfo(0, 0, out typeInfo);
            Assert.Equal(HRESULT.S_OK, hr);
            Assert.NotNull(typeInfo);
        }

        [WinFormsFact]
        public unsafe void IDispatch_GetTypeInfoCount_Invoke_Success()
        {
            using var image = new Bitmap(16, 16);
            IPictureDisp picture = SubAxHost.GetIPictureDispFromPicture(image);
            IDispatch dispatch = (IDispatch)picture;

            uint ctInfo = uint.MaxValue;
            HRESULT hr = dispatch.GetTypeInfoCount(&ctInfo);
            Assert.Equal(HRESULT.S_OK, hr);
            Assert.Equal(1u, ctInfo);
        }

        [WinFormsFact]
        public unsafe void IDispatch_Invoke_Invoke_Success()
        {
            using var image = new Bitmap(16, 32);
            IPictureDisp picture = SubAxHost.GetIPictureDispFromPicture(image);
            IDispatch dispatch = (IDispatch)picture;

            Guid riid = Guid.Empty;
            var dispParams = new DISPPARAMS();
            var varResult = new object[1];
            var excepInfo = new EXCEPINFO();
            uint argErr = 0;
            HRESULT hr = dispatch.Invoke(
                (DispatchID)4,
                &riid,
                0,
                DISPATCH.PROPERTYGET,
                &dispParams,
                varResult,
                &excepInfo,
                &argErr
            );
        }

        private class SubAxHost : AxHost
        {
            private SubAxHost(string clsidString) : base(clsidString)
            {
            }

            public new static IPictureDisp GetIPictureDispFromPicture(Image image) => (IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
        }
    }
}
