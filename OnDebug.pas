unit OnDebug;

interface

uses
  Windows, SysUtils, ToolsAPI, Dialogs, Clipbrd;

procedure Register;

type
  TDebuggerNotifier = class(TNotifierObject, IOTADebuggerNotifier)
  private
    procedure ProcessCreated(const Process: IOTAProcess); overload;
    procedure ProcessDestroyed(const Process: IOTAProcess); overload;
    procedure BreakpointAdded(const Breakpoint: IOTABreakpoint); overload;
    procedure BreakpointDeleted(const Breakpoint: IOTABreakpoint); overload;
  end;

implementation

uses
  GlobalUnit;

var
  IDENo2: Integer;

//*****************************************************************************
//登録
//*****************************************************************************
procedure Register;
begin
  IDENo2 := (BorlandIDEServices as IOTADebuggerServices).AddNotifier(TDebuggerNotifier.Create);
end;

//*****************************************************************************
//削除
//*****************************************************************************
procedure UnRegister;
begin
  (BorlandIDEServices as IOTADebuggerServices).RemoveNotifier(IDENo2);
end;


//////////////////////////////////////////////////////////////////////////
{ TDebuggerNotifier }
//////////////////////////////////////////////////////////////////////////
//*****************************************************************************
//[イベント] ProcessCreated時
//[ 概  要 ] 開いているソースで変更中のものについて、プロジェクトフォルダ配下に
//           Backupフォルダを作成し、ファイルを保存
//*****************************************************************************
procedure TDebuggerNotifier.ProcessCreated(const Process: IOTAProcess);
var
  i, j: integer;
  BackupFolder: string;
  Project: IOTAProject;
  ModuleInfo: IOTAModuleInfo;
  Module: IOTAModule;
  Editor: IOTAEditor;
  SourceEditor: IOTASourceEditor;
  EditBuffer: IOTAEditBuffer;
begin
  if not CheckEditView() then Exit;

  Project := GetCurrentProject();
  if not Assigned(Project) then Exit;
  BackupFolder := ExtractFileDir(Project.FileName) + '\Backup';

  //Moduleの数だけループ
  for i := 0 to Project.GetModuleCount - 1 do
  begin
    ModuleInfo := Project.GetModule(i);

    //フォームをエディタで表示している時、フォーム表示に切替える
    if ModuleInfo.FormName <> '' then
    begin
      Module := ModuleInfo.OpenModule;
    end
    else
    begin
      Module := GetModuleFromFileName(ModuleInfo.FileName);
    end;
    if not Assigned(Module) then Continue;

    //Editorの数だけループ
    for j := 0 to Module.ModuleFileCount - 1 do
    begin
      Editor := Module.ModuleFileEditors[j];
      if Supports(Editor, IOTASourceEditor, SourceEditor) and
         Supports(Editor, IOTAEditBuffer, EditBuffer) then
      begin
        //変更ありか？
        if SourceEditor.Modified then
        begin
          BackupSource(BackupFolder, EditBuffer);
        end;
      end;
    end;
  end;
end;

//*****************************************************************************
//[ 概  要 ] 定義がないとエラー
//*****************************************************************************
procedure TDebuggerNotifier.ProcessDestroyed(const Process: IOTAProcess); begin end;
procedure TDebuggerNotifier.BreakpointAdded(const Breakpoint: IOTABreakpoint); begin end;
procedure TDebuggerNotifier.BreakpointDeleted(const Breakpoint: IOTABreakpoint); begin end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
initialization
finalization
	UnRegister;
end.
