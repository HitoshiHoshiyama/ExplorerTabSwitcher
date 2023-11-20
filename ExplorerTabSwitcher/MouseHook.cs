using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Automation;

using NLog;

namespace ExplorerTabSwitcher
{
    /// <summary>マウスホイールメッセージをフックするクラス。</summary>
    internal class MouseHook : IDisposable
    {
        #region NativeApi block
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr WindowFromPoint(POINT point);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public System.IntPtr dwExtraInfo;
        }

        private const int WH_MOUSE_LL = 0x000E;
        private const int WM_MOUSEWHEEL = 0x020A;
        #endregion NativeApi block

        /// <summary>
        /// コンストラクタ。
        /// </summary>
        /// <param name="logger">NLogのロガーインスタンスを指定する。</param>
        /// <exception cref="Exception">プロセスのモジュールを取得できない場合に発生する。</exception>
        public MouseHook(Logger logger)
        {
            this.logger = logger;
            this.hookProcPointer = this.HookProcedure;

            using (var currentProcess = Process.GetCurrentProcess())
            using (var currentModule = currentProcess.MainModule)
            {
                if (currentModule is null) throw new Exception("MainModule is null.");
                else
                {
                    this.hookHandle = SetWindowsHookEx(WH_MOUSE_LL, this.hookProcPointer, GetModuleHandle(currentModule.ModuleName), 0);
                    this.logger.Info($"Mouse hook installed(0x{this.hookHandle}).");
                }
            }

            this.switcher = new TabSwitcher(this.HookQueue, this.cancellation, this.logger);
            this.SwitchTask = Task.Run(() => { this.switcher.SwitchTaskProc(); });
        }

        /// <summary>フック解除後、リソースを破棄する。</summary>
        public void Dispose()
        {
            var unhook = UnhookWindowsHookEx((IntPtr)this.hookHandle);
            this.logger.Info($"Unhook {(unhook ? "succeeded" : "failed")}.");

            this.cancellation.Cancel();
            this.logger.Info("Task cancel request.");
            this.SwitchTask.Wait();

            this.SwitchTask.Dispose();
            this.cancellation.Dispose();
            this.logger.Info("MuseHook all resource disposed.");
        }

        /// <summary>
        /// <br>メッセージフックを処理するコールバックプロシジャ。</br>
        /// <br>WM_MOUSEWHEELをフックしてパラメータをワーカースレッドへキューイングする。</br>
        /// </summary>
        /// <param name="nCode">フックコードが設定される。</param>
        /// <param name="wParam">メッセージコードが設定される。</param>
        /// <param name="lParam">MSLLHOOKSTRUCT構造体でパラメータが設定される。</param>
        /// <returns>フックチェーンの戻り値を設定する。</returns>
        private IntPtr HookProcedure(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEWHEEL && lParam != IntPtr.Zero)
            {
#pragma warning disable CS8605 // null の可能性がある値をボックス化解除しています。<- lParam != IntPtr.Zero でチェックしてるのにうるさいので抑制
                MSLLHOOKSTRUCT hookStr = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
#pragma warning restore CS8605 // null の可能性がある値をボックス化解除しています。
                var delta = (int)hookStr.mouseData >> 16;
                this.HookQueue.Add(new Tuple<int, int, int>(hookStr.pt.X, hookStr.pt.Y, delta));
                this.logger.Debug($"Queue add({hookStr.pt.X}, {hookStr.pt.Y} delta:{delta})");
            }

            return CallNextHookEx(this.hookHandle, nCode, wParam, lParam);
        }

        /// <summary>
        /// フックプロシジャのデリゲート。
        /// </summary>
        /// <param name="nCode">フックコードが設定される。</param>
        /// <param name="wParam">メッセージコードが設定される。</param>
        /// <param name="lParam">MSLLHOOKSTRUCT構造体でパラメータが設定される。</param>
        /// <returns>フックチェーンの戻り値を設定する。</returns>
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>タブ切り替え処理クラスのインスタンス。</summary>
        private TabSwitcher switcher;
        /// <summary>WM_MOUSEWHEELのパラメータを受け渡すキュー。</summary>
        private BlockingCollection<Tuple<int, int, int>> HookQueue = new BlockingCollection<Tuple<int, int, int>>();
        /// <summary>タブ切り替え処理のTaskハンドル。</summary>
        private Task SwitchTask;
        /// <summary>タブ切り替え処理のキャンセルオブジェクト。</summary>
        private CancellationTokenSource cancellation = new CancellationTokenSource();
        /// <summary>
        /// フックプロシジャのデリゲートインスタンス。
        /// <br>生存期間を確約するためにメンバとして保持する。</br>
        /// </summary>
        private HookProc hookProcPointer;
        /// <summary>フックハンドル。</summary>
        private IntPtr hookHandle;
        /// <summary>NLogのロガーインスタンス。</summary>
        private Logger logger;
    }

    /// <summary>タブ切り替え処理クラス。</summary>
    internal class TabSwitcher
    {
        #region NativeApi block
        [DllImport("user32.dll", EntryPoint = "MapVirtualKeyA")]
        private extern static int MapVirtualKey(int wCode, int wMapType);
        [DllImport("user32.dll", SetLastError = true)]
        private extern static uint SendInput(int nInputs, Input[] pInputs, int cbsize);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MouseInput
        {
            public int X;
            public int Y;
            public int Data;
            public int Flags;
            public int Time;
            public IntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KeyboardInput
        {
            public short VirtualKey;
            public short ScanCode;
            public int Flags;
            public int Time;
            public IntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HardwareInput
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct Input
        {
            public int Type;
            public InputUnion ui;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)]
            public MouseInput Mouse;
            [FieldOffset(0)]
            public KeyboardInput Keyboard;
            [FieldOffset(0)]
            public HardwareInput Hardware;
        }

        public const int INPUT_MOUSE = 0;
        public const int INPUT_KEYBOARD = 1;
        public const int INPUT_HARDWARE = 2;

        public const int KEYEVENTF_KEYDOWN = 0;
        public const int KEYEVENTF_KEYUP = 2;
        public const int KEYEVENTF_EXTENDEDKEY = 1;

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;

        private const int VK_TAB = 0x09;
        private const int VK_SHIFT = 0x10;
        private const int VK_CONTROL = 0x11;

        private const int MAPVK_VK_TO_VSC_EX = 4;
        #endregion NativeApi block

        /// <summary>
        /// コンストラクタ。
        /// </summary>
        /// <param name="Queue">WM_MOUSEWHEELのパラメータを受け渡すキューを指定する。</param>
        /// <param name="cancellation">処理をキャンセルするキャンセルオブジェクトを指定する。</param>
        /// <param name="logger">NLogのロガーインスタンスを指定する。</param>
        public TabSwitcher(BlockingCollection<Tuple<int, int, int>> Queue, CancellationTokenSource cancellation, Logger logger)
        {
            this.HookQueue = Queue;
            this.cancellation = cancellation;
            this.logger = logger;
        }

        /// <summary>
        /// <br>タブ切り替え処理のメインタスク。</br>
        /// <br>WM_MOUSEWHEELのパラメータがキューに登録されるまでは、処理待ち状態を維持する。</br>
        /// <br>キューから受け取ったパラメータからカーソル下のコントロールを取得し、
        /// 切り替え対象のタブであれば回転方向に応じてタブを切り替える。</br>
        /// </summary>
        public void SwitchTaskProc()
        {
            var point = new System.Windows.Point();
            int delta;
            var skipList = new List<string>();  // 対象外だったAutomationElementのRuntimeIdリスト

            while (true)
            {
                try
                {
                    var request = this.HookQueue.Take(this.cancellation.Token);

                    // Item1:X座標 Item2:Y座標 Item3:ホイールΔ(120/-120)
                    delta = request.Item3;
                    point.X = request.Item1;
                    point.Y = request.Item2;
                    // 座標から対象UI要素を取得
                    var target = AutomationElement.FromPoint(point);
                    if (target is null)
                    {
                        this.logger.Debug($"AutomationElement.FromPoint({point.X}, {point.Y}) is null.");
                        continue;
                    }
                    // RuntimeIdをキーにリストをチェックし、登録済みなら対象外なので処理をスキップ
                    var runtimeId = string.Join(", ", target.GetRuntimeId());
                    if (skipList.Contains(runtimeId))
                    {
                        this.logger.Debug($"Skip element:({runtimeId}).");
                        continue;
                    }
                    this.logger.Debug($"WM_MOUSEWHEEL({request.Item1}, {request.Item2} delta:{delta}) Target Class:{target.Current.ClassName}");

                    // 対象UI要素がタブ切り替え対象かチェック
                    var elemId = this.IdentifyElement(target, point);
                    // TODO: Ctrlキー同時押し排除
                    if (elemId.Item1 != TargetKind.Another && elemId.Item1 != TargetKind.WindowsTerminalnotTab && delta != 0)
                    {
                        // タブ切り替え処理
                        if (elemId.Item1 == TargetKind.AcrobatReader)
                        {
                            // Acrobat Reader DC(等のUIAutomationで切り替えられないタイプ)
                            var postResult = this.SendCtrlTabStroke(elemId.Item2.Current.NativeWindowHandle, delta > 0);
                            this.logger.Debug($"Tab switch to {(delta > 0 ? "left" : "right")} is {(postResult ? "succeeded" : "failed")}.");
                        }
                        else
                        {
                            // UIAutomationで切り替えられるタイプ
                            AutomationElement switchTab = (delta > 0) ? elemId.Item2 : elemId.Item3;
                            SelectionItemPattern? pattern = switchTab.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            if (pattern != null)
                            {
                                pattern.Select();
                                this.logger.Debug($"Tab switch to {switchTab.Current.Name}");
                            }
                        }
                    }
                    else if (elemId.Item1 != TargetKind.WindowsTerminalnotTab)
                    {
                        // 対象外だったRuntimeIdをリストに登録
                        runtimeId = string.Join(", ", target.GetRuntimeId());
                        if (!skipList.Contains(runtimeId))
                        {
                            skipList.Add(runtimeId);
                            this.logger.Debug($"Add to skip list:{runtimeId}");
                        }
                    }else this.logger.Debug($"Do not add to skip list:{runtimeId}");
                }
                catch (OperationCanceledException)
                {
                    // タスクキャンセル通知を受け取ったので処理を終了
                    this.logger.Info("Cancel requested.");
                    break;
                }
                catch (Exception ex)
                {
                    // 例外はログに出力してキュー待機へ戻る
                    this.logger.Warn(ex.ToString());
                }
            }
        }

        /// <summary>
        /// 指定したUI要素がタブ切り替え対象かチェックし、対象であればアクティブタブの前後タブを取得する。
        /// </summary>
        /// <param name="TargetElmArg">チェック対象のUI要素を指定する。</param>
        /// <returns><br>切り替え対象だった場合、[対象コントロールの種別, 前のタブ, 後のタブ]の順に格納されたタプルが返る。</br>
        /// <br>該当しない場合は、対象コントロールの種別にTargetKind.Anotherが設定されたタプルが返る。
        /// この場合、前後のタブに設定された内容は保証されない。</br></returns>
        private Tuple<TargetKind, AutomationElement, AutomationElement> IdentifyElement(AutomationElement TargetElmArg, System.Windows.Point? MousePoint = null)
        {
            var treeWalker = TreeWalker.ControlViewWalker;
            var TargetElm = TargetElmArg;

            if ((TargetElm.Current.ClassName == "TextBlock" || TargetElm.Current.ClassName == "Image" || TargetElm.Current.ClassName == "Button") &&
                TargetElm.Current.FrameworkId == "XAML")
            {
                // Explorerのタブ上アイコンとテキストの可能性があるので親で判定
                TargetElm = treeWalker.GetParent(TargetElm);
            }

            Tuple<TargetKind, AutomationElement, AutomationElement>? result;
            if (TargetElm.Current.ClassName == "ListViewItem" && TargetElm.Current.FrameworkId == "XAML")
            {
                // Explorerの可能性
                result = this.IdentifyExplorer(treeWalker, TargetElm);
            }
            else if (TargetElm.Current.ClassName == "TabStrip::TabDragContextImpl" && TargetElm.Current.FrameworkId == "Chrome")
            {
                // Edgeの可能性
                result = this.IdentifyEdge(treeWalker, TargetElm);
            }
            else if (TargetElm.Current.ClassName == "CASCADIA_HOSTING_WINDOW_CLASS" && TargetElm.Current.FrameworkId == "Win32")
            {
                // Windows Terminalの可能性
                result = this.IdentifyWindosTerminal(treeWalker, TargetElm, MousePoint);
            }
            else if (TargetElm.Current.ClassName == "Notepad" && TargetElm.Current.FrameworkId == "Win32")
            {
                // メモ帳の可能性(Windows Terminalとメモ帳は共通構造)
                result = this.IdentifyWindosTerminal(treeWalker, TargetElm, MousePoint);
            }
            else if (TargetElm.Current.ClassName == "AVL_AVView" && TargetElm.Current.Name == "AVTabLinksContainerViewForDocs")
            {
                // Acrobat Reader DCの可能性
                var parent = this.FindParentElement(treeWalker, TargetElm, "AcrobatSDIWindow");
                if (parent is not null) result = new Tuple<TargetKind, AutomationElement, AutomationElement>(TargetKind.AcrobatReader, parent, parent);
                else result = new Tuple<TargetKind, AutomationElement, AutomationElement>(TargetKind.Another, TargetElmArg, TargetElmArg);
            }
            else
            {
                // いずれでもない
                result = new Tuple<TargetKind, AutomationElement, AutomationElement>(TargetKind.Another, TargetElmArg, TargetElmArg);
            }

            return result;
        }

        /// <summary>
        /// 指定したUI要素がExplorerかチェックし、そうであればアクティブタブの前後タブを取得する。
        /// </summary>
        /// <param name="treeWalker">UI要素探索に使用するTreeWalkerを指定する。</param>
        /// <param name="TargetElm">チェック対象のUI要素を指定する。</param>
        /// <returns><br>切り替え対象だった場合、[対象コントロールの種別, 前のタブ, 後のタブ]の順に格納されたタプルが返る。</br>
        /// <br>該当しない場合は、対象コントロールの種別にTargetKind.Anotherが設定されたタプルが返る。
        /// この場合、前後のタブに設定された内容は保証されない。</br></returns>
        private Tuple<TargetKind, AutomationElement, AutomationElement> IdentifyExplorer(TreeWalker treeWalker, AutomationElement TargetElm)
        {
            var resultItem1 = TargetKind.Another;
            var resultItem2 = TargetElm;
            var resultItem3 = TargetElm;

            var prop = TargetElm.GetRuntimeId();
            if (prop is not null && prop.Length > 1 && prop[1] != 0)
            {
                var nativeElm = AutomationElement.FromHandle((IntPtr)prop[1]);
                if (nativeElm != null && nativeElm.Current.ClassName == "Microsoft.UI.Content.DesktopChildSiteBridge")
                {
                    // Explorerは親まで遡ってクラス名を見る必要がある
                    var parent = TreeWalker.ControlViewWalker.GetParent(nativeElm);
                    if (parent != null && parent.Current.ClassName == "CabinetWClass")
                    {
                        // たぶんExplorerで間違いないはず
                        this.logger.Debug("Explorer tab found.");
                        // ExplorerはタブそのものがZ最上位にいるようなので、親を取得する必要がある
                        var selectElms = this.GetSelectedChild(treeWalker, treeWalker.GetParent(TargetElm), "ListViewItem");
                        if (selectElms.Count > 2)
                        {
                            this.logger.Debug($"Switch tab found({selectElms[1].Current.Name} <- {selectElms[0].Current.Name} -> {selectElms[2].Current.Name}).");
                            resultItem1 = TargetKind.Explorer;
                            resultItem2 = selectElms[1];
                            resultItem3 = selectElms[2];
                        }
                    }
                }
            }
            return new Tuple<TargetKind, AutomationElement, AutomationElement>(resultItem1, resultItem2, resultItem3);
        }

        /// <summary>
        /// 指定したUI要素がEdgeかチェックし、そうであればアクティブタブの前後タブを取得する。
        /// </summary>
        /// <param name="treeWalker">UI要素探索に使用するTreeWalkerを指定する。</param>
        /// <param name="TargetElm">チェック対象のUI要素を指定する。</param>
        /// <returns><br>切り替え対象だった場合、[対象コントロールの種別, 前のタブ, 後のタブ]の順に格納されたタプルが返る。</br>
        /// <br>該当しない場合は、対象コントロールの種別にTargetKind.Anotherが設定されたタプルが返る。
        /// この場合、前後のタブに設定された内容は保証されない。</br></returns>
        private Tuple<TargetKind, AutomationElement, AutomationElement> IdentifyEdge(TreeWalker treeWalker, AutomationElement TargetElm)
        {
            var resultItem1 = TargetKind.Another;
            var resultItem2 = TargetElm;
            var resultItem3 = TargetElm;

            var prop = TargetElm.GetRuntimeId();
            if (prop is not null && prop.Length > 1 && prop[1] != 0)
            {
                var nativeElm = AutomationElement.FromHandle((IntPtr)prop[1]);
                if (nativeElm != null && nativeElm.Current.ClassName == "Chrome_WidgetWin_1")
                {
                    // たぶんEdgeで間違いないはず
                    this.logger.Debug("Edge tab found.");
                    // EdgeではUIツリーの構造上、前の兄弟要素がタブアイテムの親要素となる
                    var selectElms = this.GetSelectedChild(treeWalker, treeWalker.GetPreviousSibling(TargetElm), "EdgeTab");
                    if (selectElms.Count > 2)
                    {
                        this.logger.Debug($"Switch tab found({selectElms[1].Current.Name} <- {selectElms[0].Current.Name} -> {selectElms[2].Current.Name}).");
                        resultItem1 = TargetKind.Edge;
                        resultItem2 = selectElms[1];
                        resultItem3 = selectElms[2];
                    }
                }
            }
            return new Tuple<TargetKind, AutomationElement, AutomationElement>(resultItem1, resultItem2, resultItem3);
        }

        /// <summary>
        /// 指定したUI要素がWindows Terminal(またはメモ帳)かチェックし、そうであればアクティブタブの前後タブを取得する。
        /// </summary>
        /// <param name="treeWalker">UI要素探索に使用するTreeWalkerを指定する。</param>
        /// <param name="TargetElm">チェック対象のUI要素を指定する。</param>
        /// <returns><br>切り替え対象だった場合、[対象コントロールの種別, 前のタブ, 後のタブ]の順に格納されたタプルが返る。</br>
        /// <br>該当しない場合は、対象コントロールの種別にTargetKind.Anotherが設定されたタプルが返る。
        /// この場合、前後のタブに設定された内容は保証されない。</br></returns>
        private Tuple<TargetKind, AutomationElement, AutomationElement> IdentifyWindosTerminal(TreeWalker treeWalker, AutomationElement TargetElm, System.Windows.Point? MousePoint)
        {
            var resultItem1 = TargetKind.Another;
            var resultItem2 = TargetElm;
            var resultItem3 = TargetElm;

            var tabList = this.FindElements(TargetElm, "ListView");
            if (tabList is not null && tabList.Count > 0 && tabList[0].Current.AutomationId == "TabListView")
            {
                // たぶんWindows Terminalで間違いないはず
                var tabBox = tabList[0].Current.BoundingRectangle;
                // 座標からUI要素を取ると一番デカい領域で取れてしまうので、自力でHit判定する必要がある
                if (MousePoint is not null && (int)tabBox.Left <= MousePoint?.X && (int)tabBox.Right >= MousePoint?.X &&
                    (int)tabBox.Top <= MousePoint?.Y && (int)tabBox.Bottom >= MousePoint?.Y)
                {
                    this.logger.Debug("Windows Terminal tab found.");
                    // Windows Terminalは親ウィンドウからListViewクラスを見付けるとそこがタブアイテムの親となっている
                    var selectElms = this.GetSelectedChild(treeWalker, tabList[0], "ListViewItem");
                    if (selectElms.Count > 2)
                    {
                        this.logger.Debug($"Switch tab found({selectElms[1].Current.Name} <- {selectElms[0].Current.Name} -> {selectElms[2].Current.Name}).");
                        resultItem1 = TargetKind.WindowsTerminal;
                        resultItem2 = selectElms[1];
                        resultItem3 = selectElms[2];
                    }
                }
                else
                {
                    // Hit判定ハズレ
                    resultItem1 = TargetKind.WindowsTerminalnotTab;
                    this.logger.Debug($"Out of TabListView area({MousePoint?.X}, {MousePoint?.Y}).");
                }
            }
            return new Tuple<TargetKind, AutomationElement, AutomationElement>(resultItem1, resultItem2, resultItem3);
        }

        /// <summary>
        /// タブリストのUI要素を起点に、アクティブタブとその前後タブを取得する。
        /// </summary>
        /// <param name="treeWalker">UIツリーを探索するためのTreeWalkerを指定する。</param>
        /// <param name="element">起点となるタブリスト要素(通常、タブアイテムの親リスト)を指定する。</param>
        /// <param name="tabClassName">各タブのクラス名を指定する。</param>
        /// <returns><br>[選択タブ, 前のタブ, 後のタブ]の順に格納されたタブ要素リストを返す。</br>
        /// <br>タブが3つに満たない場合、前後タブ(タブ数:2の場合)や前後タブと選択タブ(タブ数:1)が同一要素になる。</br></returns>
        private List<AutomationElement> GetSelectedChild(TreeWalker treeWalker, AutomationElement element, string tabClassName = "ListViewItem")
        {
            var result = new List<AutomationElement>();
            var child = treeWalker.GetFirstChild(element);
            var firstChild = child as AutomationElement;
            AutomationElement? prev = null;

            // タブリスト要素から 子要素～子要素の兄弟 の順に確認していく
            while (child is not null)
            {
                if(child.Current.ClassName == tabClassName)
                {
                    SelectionItemPattern? pattern = child.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    if (pattern is not null && pattern.Current.IsSelected)
                    {
                        // 選択された(アクティブ)要素を発見
                        this.logger.Debug($"Selected tab found({child.Current.Name}).");
                        result.Add(child);
                        if (prev != null)
                        {
                            // 先頭タブ以外がアクティブ
                            this.logger.Debug($"Previous tab found({prev.Current.Name}).");
                            result.Add(prev);
                        }
                        else
                        {
                            // 先頭タブがアクティブなので、最後尾タブが前のタブ扱い(ループ)
                            var lastTab = treeWalker.GetLastChild(element);
                            this.logger.Debug($"Previous tab found({lastTab.Current.Name}).");
                            result.Add(lastTab);
                        }
                    }
                    else if (result.Count > 0)
                    {
                        // 既にアクティブタブ検出済みなので、直後のタブを後のタブとして追加して終了
                        this.logger.Debug($"Next tab found({child.Current.Name}).");
                        result.Add(child);
                        break;
                    }
                }
                prev = child;
                child = treeWalker.GetNextSibling(child);
            }
            if (result.Count == 2)
            {
                // アクティブタブ検出済みかつ後のタブ未検出(つまり右端タブがアクティブ)なので先頭タブが後のタブ扱い(ループ)
                this.logger.Debug($"Next tab is same tab({firstChild.Current.Name}).");
                result.Add(firstChild);
            }
            return result;
        }

        /// <summary>
        /// <br>親ウィンドウのAutomationElementを起点として指定クラスのAutomationElementを検索する。</br>
        /// <br>FindAllが信用ならなかったため追加。</br>
        /// </summary>
        /// <param name="rootElement">検索の起点とするAutomationElementを指定する。</param>
        /// <param name="automationClass">検索するウィンドウクラス名を指定する。</param>
        /// <param name="NeedCheckSibling">rootElementの兄弟要素を確認するか指定する。既定値は false。</param>
        /// <returns>クラス名が一致した最初のAutomationElementのリストを返す。</returns>
        private List<AutomationElement> FindElements(AutomationElement rootElement, string automationClass, bool NeedCheckSibling = false)
        {
            var result = new List<AutomationElement>();
            try
            {
                var child = TreeWalker.ContentViewWalker.GetFirstChild(rootElement);
                if (child != null)
                {
                    if (child.Current.ClassName == automationClass) result.Add(child);
                    var nextLevel = FindElements(child, automationClass, true);
                    if (nextLevel is not null) result.AddRange(nextLevel);
                }
                if (NeedCheckSibling)
                {
                    // 再帰呼び出しの時のみ通るルート
                    if (rootElement.Current.ClassName == automationClass) result.Add(rootElement);
                    var nextElement = TreeWalker.ContentViewWalker.GetNextSibling(rootElement);
                    if (nextElement != null) result.AddRange(FindElements(nextElement, automationClass, true));
                }
            }
            catch (Exception ex) { this.logger.Warn(ex.ToString()); }
            return result;
        }

        /// <summary>
        /// 指定UI要素から遡り、指定したクラスの親UI要素を取得する。
        /// </summary>
        /// <param name="treeWalker">検索に使用するTreeWalkerを指定する。</param>
        /// <param name="childElement">検索の起点とするUI要素を指定する。</param>
        /// <param name="automationClass">検索するクラス名を指定する。</param>
        /// <returns><br>発見した親UI要素を返す。</br>
        /// <br>見付からなかった場合はnullを返す。</br></returns>
        private AutomationElement? FindParentElement(TreeWalker treeWalker, AutomationElement childElement, string automationClass)
        {
            var parent = treeWalker.GetParent(childElement);
            if (parent is null) return null;
            else if (parent.Current.ClassName == automationClass) return parent;
            else return this.FindParentElement(treeWalker, parent, automationClass);
        }

        /// <summary>
        /// Ctrl+Tab(+Shift)のキーストロークを対象ウィンドウへ送信する。
        /// </summary>
        /// <param name="Hwnd">対象ウィンドウのハンドルを指定する。</param>
        /// <param name="IsReverse">Shift押下が必要な場合はtrueを指定する。</param>
        private bool SendCtrlTabStroke(IntPtr Hwnd, bool IsReverse)
        {
            var result = true;
            List<short> vkList = IsReverse ? [(short)VK_CONTROL, (short)VK_SHIFT] : [(short)VK_CONTROL];
            result &= this.SentGlobalKey(vkList, true);
            Thread.Sleep(SENDINPUT_POSTMESSAGE_WAIT_TIME);
            result &= PostMessage(Hwnd, WM_KEYDOWN, VK_TAB, 0);
            result &= PostMessage(Hwnd, WM_KEYUP, VK_TAB, 0);
            Thread.Sleep(SENDINPUT_POSTMESSAGE_WAIT_TIME);
            result &= this.SentGlobalKey(vkList, false);
            this.logger.Debug($"PostMessage to {Hwnd:x8} is {(result ? "succeeded" : "failed")}.");

            return result;
        }

        /// <summary>
        /// <br>キーストロークをエミュレーションする。</br>
        /// <br>特殊キーを合成して送信可能。</br>
        /// </summary>
        /// <param name="VirtualKeyList">仮想キーコードリストを指定する。</param>
        /// <param name="IsKeyDown">キー押下の場合はtrueを、キーを離す場合はfalseを指定する。</param>
        /// <returns>SendInputが成功した場合はtrueが返る。</returns>
        private bool SentGlobalKey(List<short> VirtualKeyList, bool IsKeyDown)
        {
            Input[] inputs = new Input[VirtualKeyList.Count];

            foreach(var (vk, idx) in VirtualKeyList.Select((val, idx)=>(val, idx)))
            {
                inputs[idx] = new Input();
                inputs[idx].Type = INPUT_KEYBOARD;
                inputs[idx].ui.Keyboard.VirtualKey = vk;
                inputs[idx].ui.Keyboard.ScanCode = (short)MapVirtualKey(vk, MAPVK_VK_TO_VSC_EX);
                inputs[idx].ui.Keyboard.Flags = IsKeyDown ? KEYEVENTF_KEYDOWN : KEYEVENTF_KEYUP;
                inputs[idx].ui.Keyboard.Time = 0;
                inputs[idx].ui.Keyboard.ExtraInfo = IntPtr.Zero;
            }

            var result = SendInput(inputs.Length, inputs, Marshal.SizeOf(inputs[0]));
            this.logger.Debug($"SendInput IsKeyDown:{IsKeyDown} inputs:{inputs.Length} result:{result}");
            return result != 0;
        }

        /// <summary>WM_MOUSEWHEELのパラメータを受け渡すキュー。</summary>
        private BlockingCollection<Tuple<int, int, int>> HookQueue = new BlockingCollection<Tuple<int, int, int>>();
        /// <summary>タスクのキャンセルオブジェクト。</summary>
        private CancellationTokenSource cancellation = new CancellationTokenSource();
        /// <summary>NLogのロガーインスタンス。</summary>
        private Logger logger;

        private const int SENDINPUT_POSTMESSAGE_WAIT_TIME = 100;
    }

    /// <summary>判定領域の結果列挙体。</summary>
    internal enum TargetKind : int
    {
        /// <summary>切り替え対象(Explorer)</summary>
        Explorer,
        /// <summary>切り替え対象(Edge)</summary>
        Edge,
        /// <summary>切り替え対象(Windows Terminal)</summary>
        WindowsTerminal,
        /// <summary>非切り替え対象(Windows Terminal)</summary>
        WindowsTerminalnotTab,
        /// <summary>切り替え対象(Acrobat Reader DC)</summary>
        AcrobatReader,
        /// <summary>非切り替え対象</summary>
        Another
    }
}
