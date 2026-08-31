using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
class ScreenCap{
[DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr d, uint f);
[DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] static extern int GetSystemMetrics(int i);
[DllImport("user32.dll")] static extern bool SetProcessDPIAware();
[StructLayout(LayoutKind.Sequential)] struct RECT{public int L,T,R,B;}
static int Main(string[] a){
SetProcessDPIAware();
try{
string mode=a[0], fn=a[1];
if(mode=="window"){
IntPtr hw=GetForegroundWindow(); RECT r; GetWindowRect(hw,out r);
int w=r.R-r.L, hh=r.B-r.T;
using(Bitmap b=new Bitmap(w,hh)){ using(Graphics g=Graphics.FromImage(b)){ IntPtr dc=g.GetHdc(); PrintWindow(hw,dc,0); g.ReleaseHdc(dc);} b.Save(fn,ImageFormat.Png);}
return 0;}
int sw=GetSystemMetrics(0), sh=GetSystemMetrics(1);
using(Bitmap b=new Bitmap(sw,sh)){
using(Graphics g=Graphics.FromImage(b)){ g.CopyFromScreen(0,0,0,0,new Size(sw,sh));}
if(mode=="region"){ int x=int.Parse(a[2]),y=int.Parse(a[3]),w=int.Parse(a[4]),h=int.Parse(a[5]);
int cx=System.Math.Max(0,x), cy=System.Math.Max(0,y), cw=System.Math.Min(w,sw-cx), ch=System.Math.Min(h,sh-cy); if(cw>0&&ch>0){ using(Bitmap c=b.Clone(new Rectangle(cx,cy,cw,ch),b.PixelFormat)){ c.Save(fn,ImageFormat.Png);} } return 0;}
b.Save(fn,ImageFormat.Png);}
return 0;
}catch(System.Exception e){ return 1;}
}}
