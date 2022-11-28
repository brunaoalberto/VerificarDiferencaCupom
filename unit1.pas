unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  DateTimePicker, AdvLed, DateUtils,
  Unit2;

type

  { TForm1 }

  TForm1 = class(TForm)
    AdvLed1: TAdvLed;
    btnCancelar: TButton;
    Button1: TButton;
    btnTotalizar: TButton;
    ckData: TCheckBox;
    dtInicio: TDateTimePicker;
    dtFim: TDateTimePicker;
    Label1: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    lbProgressBar: TLabel;
    lbTotalArquivos: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    lbTotalBruto: TLabel;
    lbTotalAcrescimos: TLabel;
    lbTotalDesconto: TLabel;
    lbTotalCancelado: TLabel;
    lbTotalLiquido: TLabel;
    ProgressBar1: TProgressBar;
    SelectDirectoryDialog1: TSelectDirectoryDialog;
    procedure btnCancelarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnTotalizarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);

  private
    FRodando: Boolean;
    procedure desativaComponentes(Sender: TObject);

  public

  end;

var
  Form1: TForm1;
  path: String;
  objeto: TVerificarXmlsPasta;

implementation

{$R *.lfm}

{ TForm1 }


function ContaArquivo(Pasta, Nome:string):integer;
var
   Rec : TSearchRec;
   Procura, Contador:integer;
begin
   Contador:=0;
   procura :=SysUtils.FindFirst(Pasta+Nome,faAnyFile, Rec);
   Contador:=0;
   while procura = 0 do
      begin
      inc(Contador);
      procura:=SysUtils.FindNext(Rec)
      end;
   SysUtils.FindClose(Rec);
   Result:=Contador;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
 qtdArquivos: Integer;
begin
  if SelectDirectoryDialog1.Execute then
    path        := SelectDirectoryDialog1.FileName + '\';
    qtdArquivos := ContaArquivo(path, 'AD*');
    if qtdArquivos > 0 then begin
      btnTotalizar.Enabled:= True;
      dtInicio.Enabled    := True;
      dtFim.Enabled       := True;
      ckData.Enabled      := True;
      AdvLed1.Kind        := lkGreenLight;
    end else begin
      btnTotalizar.Enabled:= False;
      dtInicio.Enabled    := False;
      dtFim.Enabled       := False;
      ckData.Enabled      := False;
      AdvLed1.Kind        := lkRedLight;
    end;
    lbTotalArquivos.Caption:= IntToStr(qtdArquivos) + ' Arquivos';
end;

procedure TForm1.btnCancelarClick(Sender: TObject);
begin
  objeto.Terminate;
end;

procedure TForm1.btnTotalizarClick(Sender: TObject);
begin
    desativaComponentes(Sender);
    objeto:= TVerificarXmlsPasta.Create(true);
    objeto.FreeOnTerminate    := True;
    objeto.Path               := path;
    objeto.chkBoxData         := ckData;
    objeto.dataInicio         := dtInicio;
    objeto.dataFim            := dtFim;
    objeto.lbProgresso        := lbProgressBar;
    objeto.barraProgresso     := ProgressBar1;
    objeto.labelTotalBruto    := lbTotalBruto;
    objeto.labelTotalDesconto := lbTotalDesconto;
    objeto.labelTotalCancelado:= lbTotalCancelado;
    objeto.labelTotalAcrescimo:= lbTotalAcrescimos;
    objeto.labelTotalLiquido  := lbTotalLiquido;
    objeto.Start;
    objeto.OnTerminate:= @desativaComponentes;
end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if objeto = nil then begin
     CloseAction:= caFree;
  end else if objeto.Path <> '' then begin
    if MessageDlg('Aviso', 'Existe um carregamento em andamento, deseja cancelar?',
                  mtConfirmation, [mbYes,mbNo], 0) = mrYes then
    begin
      objeto.Terminate;
      Application.Terminate;
    end else begin
      CloseAction:= caNone;
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  AdvLed1.State:= lsOn;
  dtInicio.Date:= StartOfTheMonth(IncMonth(Date, -1));
  dtFim.Date   := EndOfTheMonth  (IncMonth(Date, -1));
end;

procedure TForm1.desativaComponentes(Sender: TObject);
begin
  if btnCancelar.Visible then begin
    btnCancelar.Visible:= False;
  end else begin
    btnCancelar.Visible:= True;
  end;
  if AdvLed1.Kind = lkYellowLight then begin
    AdvLed1.Kind:= lkGreenLight;
  end else begin
    AdvLed1.Kind:= lkYellowLight
  end;
  if FRodando then begin
    FRodando:= False;
  end else begin
    FRodando:= True;
  end;
  if btnTotalizar.Enabled then begin
    btnTotalizar.Enabled:= False;
  end else begin
    btnTotalizar.Enabled:= True;
  end;
  if Button1.Enabled then begin
    Button1.Enabled:= False;
  end else begin
    Button1.Enabled:= True;
  end;
  if dtInicio.Enabled then begin
    dtInicio.Enabled:= False;
  end else begin
    dtInicio.Enabled:= True;
  end;
  if dtFim.Enabled then begin
    dtFim.Enabled:= False;
  end else begin
    dtFim.Enabled:= True;
  end;
  if ckData.Enabled then begin
    ckData.Enabled:= False;
  end else begin
    ckData.Enabled:= True;
  end;
end;

end.

