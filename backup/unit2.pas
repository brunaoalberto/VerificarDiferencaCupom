unit Unit2;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Dialogs, StdCtrls, ComCtrls, DateTimePicker, Graphics,
  ACBrSAT;

type

  { TVerificarXmlsPasta }

  TVerificarXmlsPasta = Class(TThread)
  private
    FPath: String;
    FMsg: String;
    FLbProgresso: TLabel;
    FBarraProgresso: TProgressBar;
    FChkBoxData: TCheckBox;
    FDataInicio: TDateTimePicker;
    FDataFim: TDateTimePicker;
    ACBrSAT: TACBrSAT;
    FlabelTotalBruto: TLabel;
    FlabelTotalLiquido: TLabel;
    FlabelTotalDesconto: TLabel;
    FlabelTotalCancelado: TLabel;
    FlabelTotalAcrescimo: TLabel;

    function getBarraProgresso: TProgressBar;
    function getChkBoxData: TCheckBox;
    function getDataFim: TDateTimePicker;
    function getDataInicio: TDateTimePicker;
    function getLabelTotalAcrescimo: TLabel;
    function getLabelTotalBruto: TLabel;
    function getLabelTotalCancelado: TLabel;
    function getLabelTotalDesconto: TLabel;
    function getLabelTotalLiquido: TLabel;
    function getLbProgresso: TLabel;
    function getPath: String;
    procedure setBarraProgresso(AValue: TProgressBar);
    procedure setChkBoxData(AValue: TCheckBox);
    procedure setDataFim(AValue: TDateTimePicker);
    procedure setDataInicio(AValue: TDateTimePicker);
    procedure setLabelTotalAcrescimo(AValue: TLabel);
    procedure setLabelTotalBruto(AValue: TLabel);
    procedure setLabelTotalCancelado(AValue: TLabel);
    procedure setLabelTotalDesconto(AValue: TLabel);
    procedure setLabelTotalLiquido(AValue: TLabel);
    procedure setLbProgresso(AValue: TLabel);
    procedure setPath(AValue: String);

    procedure mensagem;

    procedure atualizaTela;

  public
    property Path: String read getPath write setPath;
    property lbProgresso: TLabel read getLbProgresso write setLbProgresso;
    property barraProgresso: TProgressBar read getBarraProgresso write setBarraProgresso;
    property chkBoxData: TCheckBox read getChkBoxData write setChkBoxData;
    property dataInicio: TDateTimePicker read getDataInicio write setDataInicio;
    property dataFim: TDateTimePicker read getDataFim write setDataFim;
    property labelTotalBruto: TLabel read getLabelTotalBruto write setLabelTotalBruto;
    property labelTotalLiquido: TLabel read getLabelTotalLiquido write setLabelTotalLiquido;
    property labelTotalDesconto: TLabel read getLabelTotalDesconto write setLabelTotalDesconto;
    property labelTotalCancelado: TLabel read getLabelTotalCancelado write setLabelTotalCancelado;
    property labelTotalAcrescimo: TLabel read getLabelTotalAcrescimo write setLabelTotalAcrescimo;

    procedure Execute; override;

  end;
var
   vlrTotalBruto    : Currency;
   vlrTotalLiquido  : Currency;
   vlrTotalDesconto : Currency;
   vlrTotalCancelado: Currency;
   vlrTotalAcrescimo: Currency;

implementation

{ TVerificarXmlsPasta }

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


function TVerificarXmlsPasta.getPath: String;
begin
  Result:= FPath;
end;

procedure TVerificarXmlsPasta.setBarraProgresso(AValue: TProgressBar);
begin
  FBarraProgresso:= AValue;
end;

procedure TVerificarXmlsPasta.setChkBoxData(AValue: TCheckBox);
begin
  FChkBoxData:= AValue;
end;

procedure TVerificarXmlsPasta.setDataFim(AValue: TDateTimePicker);
begin
  FDataFim:= AValue;
end;

procedure TVerificarXmlsPasta.setDataInicio(AValue: TDateTimePicker);
begin
  FDataInicio:= AValue;
end;

procedure TVerificarXmlsPasta.setLabelTotalAcrescimo(AValue: TLabel);
begin
  FlabelTotalAcrescimo:= AValue;
end;

procedure TVerificarXmlsPasta.setLabelTotalBruto(AValue: TLabel);
begin
  FlabelTotalBruto:= AValue;
end;

procedure TVerificarXmlsPasta.setLabelTotalCancelado(AValue: TLabel);
begin
  FlabelTotalCancelado:= AValue;
end;

procedure TVerificarXmlsPasta.setLabelTotalDesconto(AValue: TLabel);
begin
  FlabelTotalDesconto:= AValue;
end;

procedure TVerificarXmlsPasta.setLabelTotalLiquido(AValue: TLabel);
begin
  FlabelTotalLiquido:= AValue;
end;

function TVerificarXmlsPasta.getLbProgresso: TLabel;
begin
  Result:= FlbProgresso;
end;

function TVerificarXmlsPasta.getBarraProgresso: TProgressBar;
begin
  Result:= FBarraProgresso;
end;

function TVerificarXmlsPasta.getChkBoxData: TCheckBox;
begin
  Result:= FChkBoxData;
end;

function TVerificarXmlsPasta.getDataFim: TDateTimePicker;
begin
  Result:= FDataFim;
end;

function TVerificarXmlsPasta.getDataInicio: TDateTimePicker;
begin
  Result:= FDataInicio;
end;

function TVerificarXmlsPasta.getLabelTotalAcrescimo: TLabel;
begin
  result:= FlabelTotalAcrescimo;
end;

function TVerificarXmlsPasta.getLabelTotalBruto: TLabel;
begin
  Result:= FlabelTotalBruto;
end;

function TVerificarXmlsPasta.getLabelTotalCancelado: TLabel;
begin
  Result:= FlabelTotalCancelado;
end;

function TVerificarXmlsPasta.getLabelTotalDesconto: TLabel;
begin
  Result:= FlabelTotalDesconto;
end;

function TVerificarXmlsPasta.getLabelTotalLiquido: TLabel;
begin
  Result:= FlabelTotalLiquido;
end;

procedure TVerificarXmlsPasta.setLbProgresso(AValue: TLabel);
begin
  FlbProgresso:= AValue;
end;

procedure TVerificarXmlsPasta.setPath(AValue: String);
begin
  FPath:= AValue;
end;

procedure TVerificarXmlsPasta.mensagem;
begin
  ShowMessage(FMsg);
end;

procedure TVerificarXmlsPasta.atualizaTela;
begin
  FlabelTotalBruto.Caption      := FormatCurr('#,0.00', vlrTotalBruto);
  FlabelTotalDesconto.Caption   := FormatCurr('#,0.00', vlrTotalDesconto);
  FlabelTotalCancelado.Caption  := FormatCurr('#,0.00', vlrTotalCancelado);
  FlabelTotalLiquido.Caption    := FormatCurr('#,0.00', vlrTotalLiquido);
  FlabelTotalAcrescimo.Caption  := FormatCurr('#,0.00', vlrTotalAcrescimo);
end;

procedure TVerificarXmlsPasta.Execute;
var
  F: TSearchRec;
  Ret: Integer;
  i: Integer;
begin
  FreeOnTerminate:= True;
  if FPath = '' then begin
    FMsg:='Escolha a pasta com os XML''s!';
    Synchronize( @mensagem );
  end else begin
    FlbProgresso.Visible     := False;
    Ret := FindFirst(FPath + '\*.xml', faAnyFile, F);
    try
      vlrTotalBruto    := 0;
      vlrTotalLiquido  := 0;
      vlrTotalDesconto := 0;
      vlrTotalCancelado:= 0;
      vlrTotalAcrescimo:= 0;

      FBarraProgresso.Max:= ContaArquivo(FPath, 'AD*');
      FBarraProgresso.Min:= 0;
      FBarraProgresso.Position:= 0;

      i:= 0;
      while Ret = 0 do begin
        ACBrSAT:= TACBrSAT.Create(nil);
        try
          ACBrSAT.CFe.LoadFromFile( FPath + F.Name );

          if FChkBoxData.Checked then begin
            if ACBrSAT.CFe.ide.dEmi >= FDataInicio.Date then begin
              if ACBrSAT.CFe.ide.dEmi < FDataFim.Date + 1 then begin
                if Pos('ADC', F.Name) > 0 then begin
                  vlrTotalCancelado:= vlrTotalCancelado
                                      + ACBrSAT.CFe.Total.vCFe;
                end;
               vlrTotalBruto   := vlrTotalBruto
                                  + ACBrSAT.CFe.Total.ICMSTot.vProd;
               vlrTotalAcrescimo:= vlrTotalAcrescimo
                                  + ACBrSAT.CFe.Total.ICMSTot.vOutro
                                  + ACBrSAT.CFe.Total.DescAcrEntr.vAcresSubtot;
               vlrTotalDesconto := vlrTotalDesconto
                                  + ACBrSAT.CFe.Total.ICMSTot.vDesc
                                  + ACBrSAT.CFe.Total.DescAcrEntr.vDescSubtot;
              end;
            end;
          end else begin
            if Pos('ADC', F.Name) > 0 then begin
              vlrTotalCancelado:= vlrTotalCancelado
                                  + ACBrSAT.CFe.Total.vCFe;
            end;
            vlrTotalBruto    := vlrTotalBruto
                                + ACBrSAT.CFe.Total.ICMSTot.vProd;
            vlrTotalAcrescimo:= vlrTotalAcrescimo
                                + ACBrSAT.CFe.Total.ICMSTot.vOutro
                                + ACBrSAT.CFe.Total.DescAcrEntr.vAcresSubtot;
            vlrTotalDesconto:= vlrTotalDesconto
                               + ACBrSAT.CFe.Total.ICMSTot.vDesc
                               + ACBrSAT.CFe.Total.DescAcrEntr.vDescSubtot;
          end;
          if (vlrTotalBruto = 0) then begin

            vlrTotalLiquido:= 0;
          end else begin
            vlrTotalLiquido:= vlrTotalBruto - vlrTotalCancelado - vlrTotalDesconto + vlrTotalAcrescimo;
          end;
          Ret := FindNext(F);
          ACBrSAT.CFe.Clear;
          if (i mod 5) = 0 then begin
             Synchronize(@atualizaTela);
             FBarraProgresso.Position:= i;
          end;
          i := i + 1;
        finally
          ACBrSAT.Free;
        end;
        if Terminated then begin
           break;
        end;
      end;
      if Not Terminated then begin
        Synchronize(@atualizaTela);
        FBarraProgresso.Position   := FBarraProgresso.Max;
        if FBarraProgresso.Position = FBarraProgresso.Max then begin
          FLbProgresso.Parent      := FBarraProgresso;
          FLbProgresso.Caption     := 'Carregamento completo!';
          FLbProgresso.AutoSize    := False;
          FLbProgresso.Transparent := True;
          FLbProgresso.Top         :=  0;
          FLbProgresso.Left        :=  0;
          FLbProgresso.Width       := FBarraProgresso.ClientWidth;
          FLbProgresso.Height      := FBarraProgresso.ClientHeight;
          FLbProgresso.Alignment   := taCenter;
          FLbProgresso.Layout      := tlCenter;
          FLbProgresso.Font.Color  := clWhite;
        end;
      end else begin
        Synchronize(@atualizaTela);
        FBarraProgresso.Position := FBarraProgresso.Max;
        FLbProgresso.Parent      := FBarraProgresso;
        FLbProgresso.Caption     := 'Processo cancelado!';
        FLbProgresso.AutoSize    := False;
        FLbProgresso.Transparent := True;
        FLbProgresso.Top         :=  0;
        FLbProgresso.Left        :=  0;
        FLbProgresso.Width       := FBarraProgresso.ClientWidth;
        FLbProgresso.Height      := FBarraProgresso.ClientHeight;
        FLbProgresso.Alignment   := taCenter;
        FLbProgresso.Layout      := tlCenter;
        FLbProgresso.Font.Color  := clRed;
      end;
    finally
      FindClose(F);
      FLbProgresso.Visible:= True;
    end;
  end;
end;

end.

