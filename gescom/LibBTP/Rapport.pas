unit Rapport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, Mask, Hctrls, HSysMenu, HTB97,Math;

type
  TfRapport = class(TFVierge)
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    Tpsfic: THCritMaskEdit;
    TpsTrait: THCritMaskEdit;
    TpsEcr: THCritMaskEdit;
    HLabel7: THLabel;
    TpsTotal: THCritMaskEdit;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  fRapport: TfRapport;

procedure AfficheRapport (FicTime,TraitTime,WriteTime,TpsTotalise : double);
implementation

{$R *.DFM}
procedure GetInfotps (Tps : double; var dd : word; var hh :word; var mm: word; var ss: word; var ms : word);
var reste : double;
begin
  dd := trunc (Tps / 86400000);
  reste := (tps - (dd * 86400000));

  hh := trunc(reste / 3600000);

  reste := (Tps - (hh*3600000));
  mm := trunc (reste / 60000);

  reste := (reste - (mm * 60000));
  ss := trunc (reste / 1000);
  
  ms := trunc(reste - (ss * 1000));
end;

procedure AfficheRapport (FicTime,TraitTime,WriteTime,TpsTotalise : double);
var XX : TFRapport;
    MyTime : TDateTime;
    dd,hh,mm,ss,ms,reste : word;
    FormatDeLaDate : string;
begin
  FormatDeLaDate := 'h"h"n"mn"s"s"';
  XX := TFRapport.create(application);
  with XX do
  begin
    GetInfotps (TpsTotalise,dd,hh,mm,ss,ms);
    MyTime := EncodeTime (hh,mm,ss,ms);
    TpsTotal.Text := formatDateTime (FormatDeLaDate,Mytime);

    GetInfotps (WriteTime,dd,hh,mm,ss,ms);
    MyTime := EncodeTime (hh,mm,ss,ms);
    TpsEcr.Text := formatDateTime (FormatDeLaDate,Mytime);

    GetInfotps (FicTime,dd,hh,mm,ss,ms);
    MyTime := EncodeTime (hh,mm,ss,ms);
    Tpsfic.Text := formatDateTime (FormatDeLaDate,Mytime);

    GetInfotps (TraitTime,dd,hh,mm,ss,ms);
    MyTime := EncodeTime (hh,mm,ss,ms);
    TpsTrait.Text := formatDateTime (FormatDeLaDate,Mytime);
  end;
  XX.ShowModal ;
  XX.Free;
end;

end.
