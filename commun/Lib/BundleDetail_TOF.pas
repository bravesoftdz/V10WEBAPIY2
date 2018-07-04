{***********UNITE*************************************************
Auteur  ...... : JSI 
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYBUNDLEDETAIL ()
Mots clefs ... : TOF;YYBUNDLEDETAIL
*****************************************************************}
Unit BundleDetail_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, FE_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF, UTob ;

function GCLanceFiche_BundleDetail(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_YYBUNDLEDETAIL = Class (TOF)

  private
    TobEltBundle, TobGEltBundle : tob;
    sCode : string;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure SetCode(LeCode : string);
    function  GetCode() : string;

    procedure ChargeElementBundle;
    procedure SetGridElement(QuelObjet : string);

    procedure AfficheObjetOnChange (Sender : TObject);

  end ;

Implementation

uses HSysMenu;

function GCLanceFiche_BundleDetail(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat = '') or (Cod = '') then exit;
  Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_YYBUNDLEDETAIL.OnArgument (S : String ) ;
begin
  Inherited ;
  SetCode(S);

  Ecran.Caption := Ecran.Caption + ' ' + S + ' ' + RechDom('YYBUNDLE', S, false);;
  UpdateCaption(Ecran);
  
  { tob contenant les éléments du bundle }
  TobEltBundle := tob.Create('',nil,-1);
  ChargeElementBundle;
  { tob permettant l'affichage }
  TobGEltBundle := tob.Create('',nil,-1);

  { Init des controles }
  if assigned(GetControl('CBOBJET')) then
    ThValComboBox(GetControl('CBOBJET')).OnChange := AfficheObjetOnChange;
  if assigned(Getcontrol('GBUNDLE')) then
  begin
    { nom de l'élément }
    ThGrid(Getcontrol('GBUNDLE')).ColWidths[1] := 50;
    { libellé }
    ThGrid(Getcontrol('GBUNDLE')).ColWidths[2] := 100;
    { type de l'objet }
    ThGrid(Getcontrol('GBUNDLE')).ColWidths[3] := 50;

    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns (ThGrid(Getcontrol('GBUNDLE')));
  end;
end ;

procedure TOF_YYBUNDLEDETAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYBUNDLEDETAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYBUNDLEDETAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYBUNDLEDETAIL.OnLoad ;
begin
  Inherited ;
  SetControlText('CBOBJET','')
  //SetGridElement('');
end ;

procedure TOF_YYBUNDLEDETAIL.OnClose ;
begin
  Inherited ;
  TobEltBundle.Free;
  TobGEltBundle.Free;
end ;

procedure TOF_YYBUNDLEDETAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYBUNDLEDETAIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYBUNDLEDETAIL.AfficheObjetOnChange (Sender : TObject);
begin
   SetGridElement(GetControlText('CBOBJET'));
end;

procedure TOF_YYBUNDLEDETAIL.SetGridElement(QuelObjet : string);
var 
    TobG : tob;
    iInd: integer;
begin
  TobGEltBundle.ClearDetail;
  if QuelObjet = '' then
  begin
    TobGEltBundle.Dupliquer(TobEltBundle,true,true);
  end else
  begin
    for iInd := 0 to TobEltBundle.Detail.Count -1 do
    begin
      if (TobEltBundle.Detail[iInd].GetString('OBJET_TYPE')= QuelObjet) then
      begin
        TobG := tob.create('',TobGEltBundle,-1);
        TobG.Dupliquer(TobEltBundle.Detail[iInd],true,true);
      end;
    end;
  end;
  TobGEltBundle.PutGridDetail(ThGrid(GetControl('GBUNDLE')),false,false,'NOM;LIBELLE;OBJET_LIB;VUE',true);
end;

procedure TOF_YYBUNDLEDETAIL.SetCode(LeCode : string);
begin
  sCode := LeCode;
end ;

function TOF_YYBUNDLEDETAIL.GetCode() : string;
begin
  result := sCode;
end ;

procedure TOF_YYBUNDLEDETAIL.ChargeElementBundle;
var jInd : integer;
    t : tob;
    sLib, sTable, sVue : string;
begin
  TobEltBundle.LoadDetailFromSql('SELECT C1.CO_ABREGE AS OBJET_TYPE,C1.CO_CODE,C1.CO_LIBELLE AS NOM,'
      + 'IIF(C1.CO_ABREGE="BUN", (SELECT C2.CO_LIBELLE FROM COMMUN C2 WHERE C2.CO_TYPE="YBU" AND C2.CO_CODE=C1.CO_LIBELLE),'
      + 'IIF(C1.CO_ABREGE="TAB",(SELECT DT_LIBELLE FROM DETABLES WHERE DT_NOMTABLE=C1.CO_LIBELLE),'
      + 'IIF(C1.CO_ABREGE="TTE",(SELECT DO_LIBELLE FROM DECOMBOS WHERE DO_COMBO=C1.CO_LIBELLE),'
      + 'IIF(C1.CO_ABREGE="VUE",(SELECT DV_LIBELLE FROM DEVUES WHERE DV_NOMVUE=C1.CO_LIBELLE),"")))) AS LIBELLE ,'
      + '(SELECT C3.CO_LIBELLE FROM COMMUN C3 WHERE C3.CO_TYPE="YOB" AND C3.CO_CODE=C1.CO_ABREGE) AS OBJET_LIB,'
      + 'C1.CO_LIBRE,C1.CO_TYPE FROM COMMUN C1 WHERE C1.CO_TYPE="YEB" AND C1.CO_LIBRE="'+sCode+'" ORDER BY C1.CO_LIBELLE');

   { gestion des vues : le CO_LIBELLE peut contenir TABLE;VUE }
   for jInd := 0 to TobEltBundle.Detail.Count -1 do
   begin
     t := TobEltBundle.Detail[jInd];
     t.AddChampSupValeur('VUE','');
     if t.GetString ('OBJET_TYPE') = 'TAB' then
     begin
       sLib := t.GetString ('NOM');
       sTable := ReadTokenSt (sLib);
       sVue := ReadTokenSt (sLib);
       if sVue <> '' then t.SetString ('VUE', sVue);
       t.SetString ('NOM', sTable);
     end;
   end;

end;

Initialization
  registerclasses ( [ TOF_YYBUNDLEDETAIL ] ) ;
end.
