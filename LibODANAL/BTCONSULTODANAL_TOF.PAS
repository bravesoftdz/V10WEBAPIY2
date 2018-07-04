{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/07/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONSULTODANAL ()
Mots clefs ... : TOF;BTCONSULTODANAL
*****************************************************************}
Unit BTCONSULTODANAL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ;

Type
  TOF_BTCONSULTODANAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBData : TOb;
    LesCols : string;
    GS : THGrid;
    //
    TheFirstMois,TheLastMois : string;
    procedure  ConstitueTOBData;
    procedure  Constituegrille;
  end ;

Implementation
uses DateUtils,UdateUtils,UtilsTOB;

procedure TOF_BTCONSULTODANAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnArgument (S : String ) ;
begin
  Inherited ;
  LesCols := 'FIXED;MOIS;BAS_STATUSGEN';
  //
  TOBData := TOB.Create('LES DATAS',nil,-1);
  TheFirstMois := LaTOB.GetString('DEPART'); if TheFirstMois = '' then TheFirstMois := '201301';
  TheLastMois := LaTOB.GetString('FIN'); if TheLastMois = '' then TheFirstMois := '201312';
  //
  GS := THGrid(GeTControl('GS'));
  ConstitueTOBData;
  Constituegrille;
end ;

procedure TOF_BTCONSULTODANAL.OnClose ;
begin
  TOBData.free;
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONSULTODANAL.ConstitueTOBData;
var Req : string;
		QQ : TQuery;
    yearD,MonthD,DayD : word;
    UneDate : TDateTime;
    CurrMois : string;
    TT : TOB;
begin
  //
  Req := 'SELECT " " AS FIXED,(SUBSTRING(BAS_MOIS,5,2)+"/"+SUBSTRING(BAS_MOIS,1,4)) AS MOIS, * '+
  			 'FROM BSTATUSANALMOIS '+
         'WHERE BAS_MOIS >= "'+TheFirstMois+'" AND BAS_MOIS < "'+TheLastMois+ '" '+
         'ORDER BY BAS_MOIS';
  QQ := OpenSQL(Req,True,1,'',true);
  if not QQ.eof then
  begin
		TOBData.LoadDetailDB('BSTATUSANALMOIS','','',QQ,false);
  end;
  ferme (QQ);
  DayD := 1;
  YearD := StrToInt(Copy(TheFirstMois,1,4));
  monthD := StrToInt(Copy(TheFirstMois,5,2));
  UneDate := EncodeDate(yearD,MonthD,Dayd);

  repeat
    CurrMois := ConstitueMois (UneDate);
    TT :=  TOBGetInto(TOBData,['BAS_MOIS'],[CurrMois]);
    if TT = nil then
    begin
			TT := TOB.Create ('BSTATUSANALMOIS',nil,-1);
      TT.AddChampSupValeur ('FIXED','');
      TT.AddChampSupValeur ('MOIS',Format('%2.2d/%4.4d',[Monthof(UneDate),Yearof(UneDate)]));
      TT.SetString('BAS_MOIS',CurrMois);
      TT.SetString('BAS_STATUSGEN','ATT');
      TOBInsertInto(TOBData,TT,['BAS_MOIS']);
    end;
    UneDate := IncMonth (UneDate,1);
  until CurrMois = TheLastMois;

end;

procedure TOF_BTCONSULTODANAL.Constituegrille;
var st      : string;
    Nam     : String;
    i       : Integer;
    depart  : Integer;
begin
  //
  GS.FixedCols := 1;
  GS.FixedRows := 1;
  //
  GS.DefaultRowHeight := 18;
  //
  GS.ColCount     := 3;
  GS.ColWidths[0] := 10;
  //
  St := LesCols;
  //
  for i := 0 to GS.ColCount - 1 do
  begin
    depart := 1;
    if i > depart then GS.ColWidths[i] := 100;
    Nam := ReadTokenSt(St);
    if Nam = 'MOIS' then
    begin
      GS.Cells[I,0] := 'Mois';
      GS.ColWidths[i] := 70;
      GS.ColAligns[i] := taCenter;
      GS.ColEditables[i] := False;
    end
    else if  (Pos(Nam ,'BAS_STATUSGEN')>0) then
    begin
      GS.Cells[I,0] := 'Status';
      GS.ColWidths[i] := 300;
      GS.ColFormats[I] := 'CB=BSTATUSGEN';
      GS.ColDrawingModes[I]:= 'IMAGE';
      GS.ColEditables[i] := False;
    end;
  end;
  //
  GS.RowCount := TOBData.detail.count + 1;
  TOBData.PutGridDetail(GS,false,false,lesCols);
end;

Initialization
  registerclasses ( [ TOF_BTCONSULTODANAL ] ) ;
end.
