{***********UNITE*************************************************
Auteur  ...... : Rohault Régis
Créé le ...... : 19/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPLISGEN ()
Mots clefs ... : TOF;CPLISGEN
*****************************************************************}
Unit CPLISGEN_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, Hdb,
     Fe_main,
     DBGrids,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,Htb97,
     HEnt1, ent1,
     UTob,
     UTOF, HQry,
     HRichEdt, Spin ,
     Buttons, Graphics,
     Grids, ExtCtrls, Windows, Registry, Menus,
     HPanel;

procedure CP_LanceFicheLisGen(St : String);

Type
  TOF_CPLISGEN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    TobGen : TOB ;
    Fiche : TFMul ;
{$IFDEF EAGLCLIENT}
    ListeGen : THGrid ;
{$ELSE}
    ListeGen : THDBGrid ;
{$ENDIF}
    QlisteGen : THQuery ;
    Pages : TpageControl ;
    PremFois : Boolean ;
    procedure InitTablesLibres(TT: TTabSheet);
    procedure FormatetafiOnClick(Sender: TObject);
    procedure MarqueOrigine;
    procedure PrechargeOrigines;
    procedure RecupMultiSel;
    procedure InitialiseLaFiche;
    procedure VireInutiles;
    procedure ChargeTobGen;
    procedure MarqueLesLignes;
    procedure BOuvrirOnCLick(Sender: TObject);
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HMsgBox;

procedure CP_LanceFicheLisGen(St : String);
begin
  AGLLanceFiche('CP','CPLISGEN','','',St);
end;

procedure TOF_CPLISGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPLISGEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPLISGEN.OnUpdate ;
begin
RecupMultiSel ;
//TobGen.saveToFile('c:\amatob\tobgen.txt',false,true,true) ;
end ;

procedure TOF_CPLISGEN.OnLoad ;
//Var TobL : Tob ;
begin
  Inherited ;
(*
If PremFois Then
  BEGIN
  If TobGen.Detail.Count>0 Then
    BEGIN
    QListeGen.First ;
    While Not QListeGen.Eof Do
      BEGIN
      TobL := TobGen.FindFirst(['G_GENERAL'],[QListeGen.FindField('G_GENERAL').AsString],False) ;
      If TobL<>NIL Then if not(ListeGen.IsCurrentSelected) then ListeGen.FlipSelection ;
      QListeGen.Next ;
      END ;
    QListeGen.First ;
    END ;
  END ;
PremFois:=FALSE ;
*)
end ;

procedure TOF_CPLISGEN.InitTablesLibres ( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : THLabel ;
    CC      : THEdit ;
BEGIN
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ; GetLibelleTableLibre('G',LesLib) ;
for i:=0 to LesLib.Count-1 do
  BEGIN
  St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
  LL:=THLabel(TForm(TT.Owner).FindComponent('TG_TABLE'+IntToStr(i))) ;
  CC:=THEdit(TForm(TT.Owner).FindComponent('G_TABLE'+IntToStr(i))) ;
  if LL<>Nil then
    BEGIN
    LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
    if ((EstSerie(S3)) and (i>2)) then LL.Visible:=False ;
    if CC<>Nil then
      BEGIN
      CC.Enabled:=LL.Enabled ; if CC.Enabled then Trouver:=True ;
      if ((EstSerie(S3)) and (i>2)) then CC.Visible:=False ;
      END ;
    END ;
  END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;

procedure TOF_CPLISGEN.OnArgument (S : String ) ;
Var St : String ;
    i : Integer ;
//    Q : TQuery ;
//    TobL : Tob ;
begin
PremFois:=TRUE ;
InitTablesLibres(TTabSheet(GetControl('PComplement'))) ;
Fiche := TFMul(Ecran) ;
ListeGen := Fiche.FListe ;
QListeGen := Fiche.Q;
Pages := TPageControl(GetControl('Pages')) ;

St:=s ;
If St<>'' Then
  BEGIN
  I:=StrToInt(St) ; TobGen:=Pointer(i) ;
{  Q:=OpenSql('SELECT G_GENERAL, G_LIBELLE FROM GENERAUX', True ) ;
  TobGen.LoadDetailDb('Gen','','',Q,True,True) ;
  Ferme(Q) ;}
  END ;
TToolBarButton97(GetControl('FormatEtafi')).OnClick:=FormatetafiOnClick ;
TToolBarButton97(GetControl('BOUVRIR')).OnClick:=BouvrirOnClick ;

If TobGen=Nil then Tobgen:=Tob.Create('Gen',nil,-1) ;
  Inherited ;
(*
If TobGen.Detail.Count>0 Then
  BEGIN
  QListeGen.Open ;
  TToolbarButton97(GetControl('BCherche')).Click ;
  QListeGen.First ;
  While Not QListeGen.Eof Do
    BEGIN
    TobL := TobGen.FindFirst(['G_GENERAL'],[QListeGen.FindField('G_GENERAL').AsString],False) ;
    If TobL<>NIL Then if not(ListeGen.IsCurrentSelected) then ListeGen.FlipSelection ;
    QListeGen.Next ;
    END ;
  QListeGen.First ;
  END ;
*)
end ;

procedure TOF_CPLISGEN.BOuvrirOnCLick(Sender : TObject) ;
{$IFDEF EAGLCLIENT}
var
  Fiche : TFMul;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
Fiche := TFMul(Ecran);
if Fiche.bSelectAll.Down then
if not Fiche.FetchLesTous then
begin
  Fiche.bSelectAllClick(nil);
  Fiche.bSelectAll.Down := False;
  Exit;
end;
{$ENDIF}
OnUpdate ;
OnClose ;
end ;

procedure TOF_CPLISGEN.FormatEtafiOnClick(Sender : TObject) ;
begin
TobGen.ClearDetail ;
InitialiseLaFiche ;
TToolbarButton97(GetControl('BCherche')).Click ;
ChargeTobGen ;
MarquelesLignes ;
end ;

procedure TOF_CPLISGEN.MarqueLesLignes ;
var i : Integer ;
    gener: string ;
begin
QListeGen.First ;
//While not QListeGen.Eof do
for i:=0 to QListeGen.recordcount-1 do
  begin
  gener:=copy(QListeGen.FindField('G_GENERAL').AsString,1,9) ;
  if ((gener>='1600') and (gener<='1629999999'))or
     ((gener>='1630') and (gener<='1639999999'))or
     ((gener>='1640') and (gener<='1649999999'))or
     ((gener>='1650') and (gener<='1659999999'))or
     ((gener>='1660') and (gener<='1689999999'))or
     ((gener>='1700') and (gener<='1999999999'))or
     ((gener>='4500') and (gener<='4539999999'))or
     ((gener>='4540') and (gener<='4549999999'))or
     ((gener>='4550') and (gener<='4599999999'))or
     ((gener>='4562') and (gener<='4562999999'))or
     ((gener>='1690') and (gener<='1699999999'))or
     ((gener>='2740') and (gener<='2749999999'))or
     ((gener>='2974') and (gener<='2974999999'))or
     ((gener>='4000') and (gener<='4039999999'))or
     ((gener>='4080') and (gener<='4083999999'))or
     ((gener>='4088') and (gener<='4089999999'))or
     ((gener>='4090') and (gener<='4095999999'))or
     ((gener>='4100') and (gener<='4179999999'))or
     ((gener>='4180') and (gener<='4189999999'))or
     ((gener>='4190') and (gener<='4195999999'))or
     ((gener>='4562') and (gener<='4562999999'))or
     ((gener>='4096') and (gener<='4099999999'))or
     ((gener>='4500') and (gener<='4539999999'))or
     ((gener>='4550') and (gener<='4561999999'))or
     ((gener>='4600') and (gener<='4639999999'))or
     ((gener>='4650') and (gener<='4669999999'))or
     ((gener>='4670') and (gener<='4679999999'))or
     ((gener>='4687') and (gener<='4699999999'))or
     ((gener>='4100') and (gener<='4189999999'))or
     ((gener>='4640') and (gener<='4649999999'))or
     ((gener>='4670') and (gener<='4679999999'))or
     ((gener>='4680') and (gener<='4686999999'))or
     ((gener>='4800') and (gener<='4859999999'))or
     ((gener>='4860') and (gener<='4869999999'))or
     ((gener>='4870') and (gener<='4899999999'))or
     ((gener>='4900') and (gener<='4909999999'))or
     ((gener>='6000') and (gener<='6029999999'))or
     ((gener>='6090') and (gener<='6092999999'))or
     ((gener>='6070') and (gener<='6089999999'))or
     ((gener>='6097') and (gener<='6097999999'))or
     ((gener>='6040') and (gener<='6049999999'))or
     ((gener>='6050') and (gener<='6059999999'))or
     ((gener>='6060') and (gener<='6069999999'))or
     ((gener>='6093') and (gener<='6096999999'))or
     ((gener>='6098') and (gener<='6099999999'))or
     ((gener>='6100') and (gener<='6119999999'))or
     ((gener>='6130') and (gener<='6139999999'))or
     ((gener>='6140') and (gener<='6149999999'))or
     ((gener>='6150') and (gener<='6159999999'))or
     ((gener>='6190') and (gener<='6199999999'))or
     ((gener>='6200') and (gener<='6219999999'))or
     ((gener>='6220') and (gener<='6229999999'))or
     ((gener>='6230') and (gener<='6239999999'))or
     ((gener>='6240') and (gener<='6249999999'))or
     ((gener>='6250') and (gener<='6259999999'))or
     ((gener>='6290') and (gener<='6299999999'))or
     ((gener>='6400') and (gener<='6449999999'))or
     ((gener>='6600') and (gener<='6659999999'))or
     ((gener>='6680') and (gener<='6699999999'))or
     ((gener>='6700') and (gener<='6749999999'))or
     ((gener>='6780') and (gener<='6799999999'))or
     ((gener>='7000') and (gener<='7011999999'))or
     ((gener>='7090') and (gener<='7091199999'))or
     ((gener>='7020') and (gener<='7021999999'))or
     ((gener>='7092') and (gener<='7092199999'))or
     ((gener>='7030') and (gener<='7031999999'))or
     ((gener>='7093') and (gener<='7093199999'))or
     ((gener>='7040') and (gener<='7041999999'))or
     ((gener>='7094') and (gener<='7094199999'))or
     ((gener>='7050') and (gener<='7051999999'))or
     ((gener>='7095') and (gener<='7095199999'))or
     ((gener>='7012') and (gener<='7019999999'))or
     ((gener>='70912') and (gener<='7091999999'))or
     ((gener>='7022') and (gener<='7029999999'))or
     ((gener>='70922') and (gener<='7092999999'))or
     ((gener>='7032') and (gener<='7039999999'))or
     ((gener>='70932') and (gener<='7093999999'))or
     ((gener>='7042') and (gener<='7049999999'))or
     ((gener>='70942') and (gener<='7094999999'))or
     ((gener>='7052') and (gener<='7059999999'))or
     ((gener>='70952') and (gener<='7095999999'))or
     ((gener>='7060') and (gener<='7061999999'))or
     ((gener>='7096') and (gener<='7096199999'))or
     ((gener>='7080') and (gener<='7081199999'))or
     ((gener>='7082') and (gener<='7082199999'))or
     ((gener>='7083') and (gener<='7083199999'))or
     ((gener>='7084') and (gener<='7084199999'))or
     ((gener>='7085') and (gener<='7085199999'))or
     ((gener>='7086') and (gener<='7086199999'))or
     ((gener>='7087') and (gener<='7087199999'))or
     ((gener>='7088') and (gener<='7088199999'))or
     ((gener>='7098') and (gener<='7099999999'))or
     ((gener>='70812') and (gener<='7081999999'))or
     ((gener>='70822') and (gener<='7082999999'))or
     ((gener>='70832') and (gener<='7083999999'))or
     ((gener>='70842') and (gener<='7084999999'))or
     ((gener>='70852') and (gener<='7085999999'))or
     ((gener>='70862') and (gener<='7086999999'))or
     ((gener>='7062') and (gener<='7069999999'))or
     ((gener>='70962') and (gener<='7096999999'))or
     ((gener>='7070') and (gener<='7071999999'))or
     ((gener>='7097') and (gener<='7097999999'))or
     ((gener>='7072') and (gener<='7079999999'))or
     ((gener>='70972') and (gener<='7097999999'))or
     ((gener>='7100') and (gener<='7199999999'))or
     ((gener>='7200') and (gener<='7299999999'))or
     ((gener>='7600') and (gener<='7619999999'))or
     ((gener>='7620') and (gener<='7629999999'))or
     ((gener>='7630') and (gener<='7659999999'))or
     ((gener>='7680') and (gener<='7699999999'))or
     ((gener>='7700') and (gener<='7749999999'))or
     ((gener>='7780') and (gener<='7799999999'))then
    begin
{$IFDEF EAGLCLIENT}
    if not(listegen.IsSelected(i)) then listegen.FlipSelection(i)
{$ELSE}
    if not(ListeGen.IsCurrentSelected) then ListeGen.FlipSelection ;
{$ENDIF}
    end else
    begin
{$IFDEF EAGLCLIENT}
    if listegen.IsSelected(i) then listegen.FlipSelection(i)
{$ELSE}
    if ListeGen.IsCurrentSelected then ListeGen.FlipSelection ;
{$ENDIF}
    end ;
  //QListeGen.Next ;
  end ;
end ;


procedure TOF_CPLISGEN.ChargeTobGen ;
var Q: TQuery ;
begin
Q:=OpenSql('SELECT G_GENERAL,G_LIBELLE FROM GENERAUX  ', True) ;
TobGen.LoadDetailDb('Gen','','',Q,True,True) ;
Ferme(Q) ;
if Tobgen.Detail.Count>0 then TobGen.detail[0].AddChampSup('Ok,',True) ;
end ;

procedure TOF_CPLISGEN.InitialiseLaFiche ;
begin
SetControlText('G_GENERAL','') ;
SetControlText('G_GENERAL_','') ;
SetControlText('G_CONSO','') ;
SetControlProperty('G_COLLECTIF','State',CbGrayed) ;
SetControlProperty('G_LETTRABLE','State',CbGrayed) ;
SetControlProperty('G_FERME','State',CbGrayed) ;
SetControlProperty('G_POINTABLE','State',CbGrayed) ;
ThValComboBox(GetControl('G_NATUREGENE')).ItemIndex:=0 ;
SetControlText('G_TABLE0','') ;
SetControlText('G_TABLE1','') ;
SetControlText('G_TABLE2','') ;
SetControlText('G_TABLE3','') ;
SetControlText('G_TABLE4','') ;
SetControlText('G_TABLE5','') ;
SetControlText('G_TABLE6','') ;
SetControlText('G_TABLE7','') ;
SetControlText('G_TABLE8','') ;
SetControlText('G_TABLE9','') ;
TToolbarButton97(GetControl('BEffaceAvance')).Click ;
end ;

procedure TOF_CPLISGEN.PrechargeOrigines ;
Var Q : TQuery ;
    St, St1 : String ;
BEGIN
TobGen.ClearDetail ;
St := 'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX ' ;
St1:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
St := St + St1 ;
Q:=OpenSQL(St, True) ;
TobGen.LoadDetailDB('Gen','','',Q,False,False) ;
Ferme(Q) ;
If TOBGen.Detail.Count>0 then TOBGen.Detail[0].AddChampSup('OK',True) ;
END ;


procedure TOF_CPLISGEN.RecupMultiSel ;
Var i :integer ;
BEGIN
PrechargeOrigines ;
If Not ListeGen.AllSelected then
   BEGIN
   for i:=0 to ListeGen.NbSelected-1 do
       BEGIN
       ListeGen.GotoLeBookmark(i) ;
       MarqueOrigine ;
       END ;
   VireInutiles ;
   END ;
END ;

Procedure TOF_CPLISGEN.VireInutiles ;
Var i : integer ;
    TobGenDetail  : TOB ;
BEGIN
for i:=TOBGen.Detail.Count-1 downto 0 do
    BEGIN
    TobGenDetail:=TOBGen.Detail[i] ;
    If TobGenDetail.GetValue('Ok')<>'X' then FreeAndNil(TobGenDetail);
    END ;
END ;

Procedure TOF_CPLISGEN.MarqueOrigine ;
Var TobGenDetail : TOB ;
BEGIN
TobGenDetail := TobGen.FindFirst(['G_GENERAL'],[QListeGen.FindField('G_GENERAL').AsString],False) ;
while TobGenDetail <> Nil do
  BEGIN
  TobGenDetail.PutValue('Ok','X') ;
  TobGenDetail := TobGen.FindNext(['G_GENERAL'],[QListeGen.FindField('G_GENERAL').AsString],False)
  END ;
END ;


procedure TOF_CPLISGEN.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPLISGEN ] ) ;
end.
