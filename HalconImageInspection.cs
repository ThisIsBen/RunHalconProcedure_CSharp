using HalconDotNet;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace アプリ
{
    class HalconImageInspection
    {
        //Declare HALCON program and procedures that we want to call
        private HDevProgram HalconProgram;
        private HDevProcedure InitProc;
        private HDevProcedure ProcessingProc;
        private HDevProcedure OutputHALCONResultImgProc;
        // instance of the engine
        private HDevEngine HalconEngine = new HDevEngine();
        // procedure calls
        private HDevProcedureCall InitProcCall;
        private HDevProcedureCall ProcessingProcCall;
        

        //HALCONで取得した特徴量をCSVに出力できる状態になっているかどうかを記録する
        private bool IsAbleToOuputHalconCSVFile = false;



        //HALCON HDevEngineの初期化＋HALCONプログラム内の初期化Procedureを実行する。
        public HalconImageInspection()
        {
            try
            {
                // enable execution of JIT compiled procedures
                HalconEngine.SetEngineAttribute("execute_procedures_jit_compiled", "true");

                // load the HALCON program
                HalconProgram = new HDevProgram(GlobalConstants.HalconHDevPath);
                // specify which HALCON procedures to call and initialize them.
                InitProc = new HDevProcedure(HalconProgram, "InitProc");
                ProcessingProc = new HDevProcedure(HalconProgram, "ProcessingProc");
                OutputHALCONResultImgProc = new HDevProcedure(HalconProgram, "OutputHALCONResultImgProc");

                // enable execution of JIT compiled procedures
                InitProc.CompileUsedProcedures();
                ProcessingProc.CompileUsedProcedures();
                OutputHALCONResultImgProc.CompileUsedProcedures();

                //create HDevProcedureCall objects
                InitProcCall = new HDevProcedureCall(InitProc);
                ProcessingProcCall = new HDevProcedureCall(ProcessingProc);
                

                //HALCONで取得した特徴量をCSVに出力できる状態になっているかどうかを判断する
                IsAbleToOuputHalconCSVFile = check_IsAbleToOuputHalconCSVFile();


                //Halconプログラムの初期化を実行する
                InitProcCall.Execute();

                //get the name of the PI対象名称
                HTuple Target = InitProcCall.GetOutputCtrlParamTuple("Target");
                GlobalConstants.Target = Target[0];

                //HALCONで取得した特徴量をCSVファイルに出力できる状態の場合、
                //HALCONで各画像から取得したCSVに出力したい特徴量の凡例（タイトル）を取得し、
                //毎回ファイルを出力する前にファイルの一行目に追記できるように保存する。
                if (IsAbleToOuputHalconCSVFile == true)
                {
                    getHalconCSVOutputTitle();
                }


            }
            //HDevEngine を使って、画像処理プログラムをロードする時、エラーが発生したら、
            //MessageBoxでオペレーターに知らせ、再起動してもらう。
            //再起動してもこのエラーが続ける場合、を停止して管理者に連絡してもらう
            catch (Exception e)
            {

                //display this error message to inform the user
                string errorMessage = "画像処理プログラムをロードする時、エラーが発生した。\n\n停止ボタンを押して、\nを再開してください。\n\n再起動してもこのエラーが続ける場合、\nを停止して管理者に連絡してください。" + "\n\nエラーメッセージ：\n" + e.Message;

                // MessageBoxでオペレーターに知らせ、再起動してもらう。
                //show the error message box to inform the user if the same message box is not being shown 
                ReportErrorMsg.showMsgBox_IfNotShown(errorMessage, " " + "App"+GlobalConstants.cameraNo+"_画像処理プログラムロードエラー");

                Console.WriteLine("画像処理プログラムをロードする時、エラーが発生した。\n\n" + "エラーメッセージ：\n" + e.Message);

                //Give up processing this picture and proceed to process the next one
                return;

            }

        }




        //対象画像をHALCONプログラム内の画像処理Procedureに入力し、検査する。
        //検査結果と画像処理エラーメッセージ（画像処理にエラーが発生した場合）を返す
        public HalconOutputModel doImageInspection(string imageToBeChecked, DateTime imageCreationTime)
        {
            //[Halcon画像処理の引数設定]
            //パラメーターを設定する。
            ProcessingProcCall.SetInputCtrlParamTuple("imageToBeChecked", imageToBeChecked);
            
            //数値トレンドデータの出力に使う画像の生成日付と時刻をHALCON画像処理プログラムに渡す。
            
            //HALCON画像処理は、数値/カテゴリートレンドデータ両方出力できるようにしたため、
            //設定ファイルで数値トレンドデータを出力しないと設定しても、
            //画像の生成日付と時刻をHALCON画像処理プログラムに渡すことが必要である。         
            ProcessingProcCall.SetInputCtrlParamTuple("imageCreationDate", imageCreationTime.ToString("yyyy/MM/dd"));
            ProcessingProcCall.SetInputCtrlParamTuple("imageCreationTime", imageCreationTime.ToString("HH:mm:ss"));


            //[HALCON画像処理を実施する]
            ProcessingProcCall.Execute();


            //[処理結果の取得]
            HTuple imageInspectionResult = ProcessingProcCall.GetOutputCtrlParamTuple("imageInspectionResult");
            //画像処理用固定名称で表示用事象名とカテゴリーデ番号を取得する。
            //カテゴリーデ番号が存在していない場合、取得したのは""である。
            InspectionDetailModel inspectionResult = InspectionDetailController.setUpInspectionResult(imageInspectionResult[0]);


            //[画像処理のエラーメッセージの取得]
            //エラーが発生していない場合、取得したのは空である。
            string PopUpErrorMsg = ProcessingProcCall.GetOutputCtrlParamTuple("PopUpErrorMsg")[0];



            //[取得した特徴量の取得]  
            string outputCSVFeatures ="";
            if (IsAbleToOuputHalconCSVFile==true)
            {
                //HALCON特徴量出力ファイルの出力ができる状態であれば、
                //出力された特徴量を取得する
                HTuple HalconCSVOutputContent = ProcessingProcCall.GetOutputCtrlParamTuple("HalconCSVOutputContent");
                if (HalconCSVOutputContent.Length > 0)
                {
                    outputCSVFeatures = HalconCSVOutputContent[0];
                }
            }


            //[現場確認用の処理結果画像の保存]
            //HALCON処理が返したWillOutputHALCONResultImage==true
            //かつ設定ファイルに現場確認用処理結果画像の保存が必要だと指定された場合のみ、
            //現場確認用処理結果画像を保存する。
            if (GlobalConstants.wantOnSiteCheckHALCONResultPic == true)
            {
                //[HALCON処理が判断した現場確認用処理結果画像保存必要性の取得]
                bool WillOutputHALCONResultImage = ProcessingProcCall.GetOutputCtrlParamTuple("WillOutputHALCONResultImage")[0];

                if (WillOutputHALCONResultImage == true)
                {
                    //Get the output HALCON image processing result Image
                    HImage HALCONResultImage = ProcessingProcCall.GetOutputIconicParamImage("HALCONResultImage");

                    //Use a ThreadPool thread to output the 処理結果画像 to NAS if the 処理結果画像 is not null.
                    if (HALCONResultImage != null)
                    {
                        Task.Run(() => outputHALCONResultImage(inspectionResult.displayJigoYocyoReasonName, imageToBeChecked, HALCONResultImage, imageCreationTime));
                    }
                }
            }


            //[画像処理結果を返す]
            return  new HalconOutputModel(inspectionResult, PopUpErrorMsg, outputCSVFeatures);
            
        }



        //HALCONで取得したCSVに出力したい特徴量があるかどうかを判断する
        private bool check_IsAbleToOuputHalconCSVFile()
        {

            //Halconプログラムの"ProcessingProc"プロシージャの
            //3つ目の戻り値の名前は"HalconCSVOutputContent"の場合のみ、
            //HALCONで取得したCSVに出力したい特徴量があると判断する。

            if( ProcessingProcCall.GetProcedure().GetOutputCtrlParamCount() > 2)
            {
                return  "HalconCSVOutputContent" == ProcessingProcCall.GetProcedure().GetOutputCtrlParamName(3);
            }
            else
            {
                return false;
            }
                
        }


        //HALCONで各画像から取得したCSVに出力したい特徴量の凡例（タイトル）を取得し、
        //毎回ファイルを出力する時にファイルの一行目に追記できるように保存する。
        private void getHalconCSVOutputTitle()
        {
            if (InitProcCall.GetProcedure().GetOutputCtrlParamCount() > 1)
            {
                if ("HalconCSVOutputTitle" == InitProcCall.GetProcedure().GetOutputCtrlParamName(2))
                {
                    HTuple HalconCSVOutputTitle = InitProcCall.GetOutputCtrlParamTuple("HalconCSVOutputTitle");
                    GlobalConstants.HalconCSVOutputTitle = HalconCSVOutputTitle[0];
                }
            }
        }




        //現場確認用処理結果画像をNASに保存する。
        private void outputHALCONResultImage(string imageInspectionResult, string imageToBeChecked, HImage HALCONResultImage, DateTime originalPicCreationTime)
        {
            //[Step1 NAS上の保存先を確保する]
            //Get the final destination folder for the current picture.
            //ケース1　フォルダーの作成が必要ない場合、デフォルトフォルダーパスを返す。
            //ケース2　フォルダーの作成が必要ない場合、
            //フォルダーの作成が成功した場合、そのフォルダーのパスを返す。
            //失敗した場合、"新しいフォルダー作成失敗"の文字を返す。
            string finalDestinationFolder = CreateFolders.createNewFolder_IfPicCreationTimeHourChange(GlobalConstants.onSiteCheckHALCONResultPicPath, originalPicCreationTime, ref GlobalConstants.lastSavedHALCONResultPicMonth, ref GlobalConstants.lastSavedHALCONResultPicDay, ref GlobalConstants.lastSavedHALCONResultPicHour);
            
            if (finalDestinationFolder == "新しいフォルダー作成失敗")
            {
                //Give up saving this picture.
                return;
            }





            //[Step2 処理結果画像の保存]
            //NAS上の保存先が確保された場合、下記の手順で処理結果画像を保存する。            
            //[Step2-1 Create procedure call for 処理結果画像の保存]
            HDevProcedureCall OutputHALCONResultImgProcCall = new HDevProcedureCall(OutputHALCONResultImgProc);


            //[Step2-2 Input parameterの設定]            
            //Input the HALCON image processing result Image
            OutputHALCONResultImgProcCall.SetInputIconicParamObject("HALCONResultImage", HALCONResultImage);
            //Input the imageInspectionResult
            OutputHALCONResultImgProcCall.SetInputCtrlParamTuple("imageInspectionResult", imageInspectionResult);
            // Input the image creation time of the imageToBeChecked
            OutputHALCONResultImgProcCall.SetInputCtrlParamTuple("ImageCreationTimestamp", originalPicCreationTime.ToString("yyyyMMdd_HHmmss.fff"));
            //Input the 処理結果画像の保存先
            OutputHALCONResultImgProcCall.SetInputCtrlParamTuple("NGImageFolderPath", finalDestinationFolder);


            //[Step2-3処理結果画像保存の実行]
            //Retry  when error occurs during compression.
            for (int retryTimes = 1; retryTimes <= GlobalConstants.retryTimesLimit; retryTimes++)
            {
                try
                {
                    //HALCONのプロシージャを呼び出して処理結果画像の保存を実行する。
                    OutputHALCONResultImgProcCall.Execute();
                    return;
                }
                
                //Retry for 3 times and output error message  
                //for the rest of the exceptions which happen while copying pictures to NAS
                catch (Exception e)
                {

                    // If it's still within retry times limit
                    if (retryTimes < GlobalConstants.retryTimesLimit)
                    {
                        //wait a while before starting next retry
                        Thread.Sleep(GlobalConstants.retryTimeInterval);
                    }

                    //If it has reached the retry limit
                    else
                    {

                        //display this error message to inform the user
                        string　errorMessage = imageToBeChecked + " の処理結果画像をNASに保存する途中でエラーが発生しました。\n今回は" + GlobalConstants.retryTimesLimit + "回目のRetryです。\nRetry回数の上限に達しましたので、\nこの画像をNASに保存するのを諦めて、次の画像を保存する。\n\n" + "エラーメッセージ：\n" + e.Message;

                        //show a pop-up message to inform the operator that if this error
                        //occurs so many times, contact 管理者 for help.
                        ReportErrorMsg.showMsgBox_IfNotShown(imageToBeChecked+" の処理結果画像を" + GlobalConstants.retryTimesLimit + "回RetryしてもNASに保存できない。\nその画像を保存するのを諦めて、次の画像を保存する。\n\nこのエラー何回も発生した場合、\n停止ボタンを押してを停止して、\n管理者に連絡してください。" + "\n\n管理者への解決手順：\nStep1 NASへの接続とNASの状態を確認してください。\nStep2 システムを再起動して、このエラーが長い時間で一回だけ発生した場合、単純に一時的なNASへの接続不具合だからです。\n多発している場合、と前後の画像をNASに保存する機能の修正が必要となります。" + "\n\nこのエラーの原因：\n" + e.Message, " " + GlobalConstants.Target + "_現場確認用画像をNASに保存できないエラー");

                        //output the error message 
                        //to the DataManagementApp_エラーメッセージ folder in NAS
                        ReportErrorMsg.outputErrorMsg("現場確認用処理結果画像をNASに保存", errorMessage);

                        //wait a while to let the program output the 3rd retry error txt message
                        //The 3rd retry error message can not be output without this wait 
                        Thread.Sleep(1000);

                    }
                }
            }


        }


    }
}
