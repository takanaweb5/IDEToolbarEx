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
//�o�^
//*****************************************************************************
procedure Register;
begin
  IDENo2 := (BorlandIDEServices as IOTADebuggerServices).AddNotifier(TDebuggerNotifier.Create);
end;

//*****************************************************************************
//�폜
//*****************************************************************************
procedure UnRegister;
begin
  (BorlandIDEServices as IOTADebuggerServices).RemoveNotifier(IDENo2);
end;


//////////////////////////////////////////////////////////////////////////
{ TDebuggerNotifier }
//////////////////////////////////////////////////////////////////////////
//*****************************************************************************
//[�C�x���g] ProcessCreated��
//[ �T  �v ] �J���Ă���\�[�X�ŕύX���̂��̂ɂ��āA�v���W�F�N�g�t�H���_�z����
//           Backup�t�H���_���쐬���A�t�@�C����ۑ�
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

  //Module�̐��������[�v
  for i := 0 to Project.GetModuleCount - 1 do
  begin
    ModuleInfo := Project.GetModule(i);

    //�t�H�[�����G�f�B�^�ŕ\�����Ă��鎞�A�t�H�[���\���ɐؑւ���
    if ModuleInfo.FormName <> '' then
    begin
      Module := ModuleInfo.OpenModule;
    end
    else
    begin
      Module := GetModuleFromFileName(ModuleInfo.FileName);
    end;
    if not Assigned(Module) then Continue;

    //Editor�̐��������[�v
    for j := 0 to Module.ModuleFileCount - 1 do
    begin
      Editor := Module.ModuleFileEditors[j];
      if Supports(Editor, IOTASourceEditor, SourceEditor) and
         Supports(Editor, IOTAEditBuffer, EditBuffer) then
      begin
        //�ύX���肩�H
        if SourceEditor.Modified then
        begin
          BackupSource(BackupFolder, EditBuffer);
        end;
      end;
    end;
  end;
end;

//*****************************************************************************
//[ �T  �v ] ��`���Ȃ��ƃG���[
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
