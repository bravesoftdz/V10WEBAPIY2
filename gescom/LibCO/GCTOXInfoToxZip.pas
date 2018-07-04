{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Consultation des informations sur le fichier TOX.ZIP
Mots clefs ... : CONSULTATION;ZIP;TOX
*****************************************************************}
unit GCTOXInfoToxZip;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, uTob, Mask, UIUtil, HPanel ;

type
  TFicheInfoToxZip = class(TForm)
    LBefore: TLabel;
    LAfter: TLabel;
    LRate: TLabel;
    LDate: TLabel;
    LTime: TLabel;
    EBfore: TLabel;
    EAfter: TLabel;
    ERate: TStaticText;
    EDate: TStaticText;
    ETime: TStaticText;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure InfoToxZip ( MaTob : tob ) ;

var
  FicheInfoToxZip: TFicheInfoToxZip;

implementation

{$R *.DFM}

procedure InfoToxZip ( MaTob : tob ) ;
var X  : TFicheInfoToxZip ;
    PP : THPanel ;
    i  : double ;
begin
PP:=FindInsidePanel ;
X:=TFicheInfoToxZip.Create(Application) ;
if PP=Nil then
  begin
  Try
    X.Caption      := X.Caption + ' ' + string ( MaTob.Detail[0].GetValue ( 'ZF_NAME' )) ;

    i :=  double ( MaTob.Detail[0].GetValue ( 'ZF_SIZEREAL' )) ;
    X.EBfore.Caption := FormatFloat ( '#,###,###,##0', i ) ;

    i :=  double ( MaTob.Detail[0].GetValue ( 'ZF_SIZECOMP' )) ;
    X.EAfter.Caption  := FormatFloat ( '#,###,###,##0', i ) ;

    i :=  100 - double ( MaTob.Detail[0].GetValue ( 'ZF_SIZECOMP' )) / double ( MaTob.Detail[0].GetValue ( 'ZF_SIZEREAL' )) * 100  ;
    X.ERate.Caption  := FormatFloat ( '#0.00', i ) + ' %' ;

    X.EDate.Caption   := string ( MaTob.Detail[0].GetValue ( 'ZF_DATE' )) ;
    X.ETime.Caption   := string ( MaTob.Detail[0].GetValue ( 'ZF_TIME' )) ;
    X.ShowModal;
  finally
    X.Free;
  end;
   end else
   begin
     InitInside(X,PP) ;
     X.Show ;
   end ;
end ;

end.
