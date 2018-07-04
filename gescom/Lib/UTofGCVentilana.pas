unit UTofGCVentilana;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
   dbTables, db,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, Vierge,UTob,Grids,EntGC,utilgc,HTB97,Tablette,
      DicoAF,graphics,Saisutil,  menus,utilarticle,M3FP;

Type
     TOF_GCVENTILANA = Class (TOF)
     public
        Faffiche, FDispo, FafficheLib, FDispoLib : TListBox;
        procedure OnUpdate ; override ;
        procedure OnDelete ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose;  override ;

     private
        TOBAffich, TOBDispo, TOBAffichLib, TOBDispoLib : TOB;
        CodeEtab , CodeAxe : String;
        SaisieEncours, bModif : Boolean;
        // MAJ Table / ecran
        procedure AskMajVentil;
        procedure MajVentil;
        procedure CreatTobStruc (TObMaj:TOB; NumItem,Cpt:integer; TraiteLib:Boolean);
        procedure ChargeVentil;
        procedure ChargeItem (TObDet: TOB; NumItem:integer; TraiteLib,Affich:Boolean);
        function TraiteChampLibre(Libelle:String):String ;

        // Gestion des boutons de déplacement
        procedure MoveItem(Sens : String);
        procedure BUp(TraiteLib : boolean);
        procedure BDown(TraiteLib : boolean);
        procedure BLeft(TraiteLib : boolean);
        procedure BRight(TraiteLib : boolean);
        function  CtrlValidite(TraiteLib : boolean) : boolean;
        procedure ClearEcran;
        function MajPosition (TraiteLib : boolean) : boolean;
        function AdapteLngChamps (Lng :integer ; Champs : string ) : integer;
        function FabriqueWhereItemDispo : String;
        function AdapteLibelleItems (wCode,wlib : string) : string;
        // Chargement Allocation
        procedure ToutAllouer;
        procedure ToutLiberer;
        procedure ChargeLesListBox ;
        procedure SaisieEnabled (bMode : Boolean);
        procedure ChargeLesTob (Traitelib,TobAffichOnly : Boolean) ;
        // Evenements
        procedure FAfficheClick(Sender: TObject);
        procedure FAfficheLibClick(Sender: TObject);
        Procedure FRecupLng(TraiteLib : boolean); // modif longueur
        Procedure FRecupsSection; // modif creat code sous section
        END ;

Type TListeCol = Class
  champs, Lib,IsLib : String;
  position , longueur : integer;
  sSection : Boolean;
END;

     const
	// libellés des messages de la TOF  afprofilgener
	TexteMsgVentil: array[1..2] of string 	=
          (
          {1}        'La longueur maximum du champs est dépassée'
          {2}       ,'La section analytique ne peut être supérieure à 17 car.'
          );

procedure AGLMoveItemVentil( parms: array of variant; nb: integer );
procedure AGLChargeVentil( parms: array of variant; nb: integer );

implementation

{ TOF_AFVENTILANA}


procedure TOF_GCVENTILANA.OnClose;
Var i : integer;
Begin
AskMajVentil;
Toutliberer;
End;

procedure TOF_GCVENTILANA.OnArgument(stArgument : String );
Var CC1,CC2,CC3,CC4 : THedit;  Combo   : THValComboBox;
BEGIN
Inherited;
// Condition plus des axes / nombre d'axes utilisés
Combo := THValComboBox (GetControl ('GSA_AXE') );
if VH_GC.GCVENTAXE2 then Combo.Plus := ' AND (X_AXE= "A1" or X_AXE= "A2")'
                    else Combo.Plus := ' AND X_AXE= "A1"';

Faffiche   := TListBox   (GetControl('FAFFICHE'));
FDispo     := TListBox   (GetControl('FDISPO'));
FafficheLib:= TListBox   (GetControl('FAFFICHELIB'));
Faffiche.OnClick := FAfficheClick; FafficheLib.OnClick := FAffichelibClick;
ToutAllouer;
ChargeLesTob(False,True);  ChargeLesTob(True,True);
ChargeLesListBox;
SaisieEnabled (False);
SetFocusControl ('GSA_AXE');
END;

procedure TOF_GCVENTILANA.OnUpdate;
Var Q : TQuery ;
BEGIN
Transactions(MajVentil,2);
// mcd 19/06/01 car si modif, table non rechargée
VH_GC.GCTOBAna.ClearDetail;
Q:=OpenSQL('SELECT * FROM STRUCTANA',True,-1,'STRUCTANA') ;
if Not Q.EOF then VH_GC.GCTOBAna.LoadDetailDB('STRUCTANA','','',Q,False) ;
Ferme(Q);
BModif:=False ;
END;

procedure TOF_GCVENTILANA.AskMajVentil;
BEGIN
 if (bModif) then
    Begin
    If (PGIAskAF('Voulez vous enregistrer la ventilation analytique','')= mrYes) then
       Begin
       Transactions(MajVentil,2);
       End;
    End;
END;

procedure TOF_GCVENTILANA.MajVentil;
Var Req : string;
    TobMaj : TOB;
    Cpt,wi : Integer;
BEGIN
    // supression des éléments existants (code + libellé)
    Req := 'DELETE FROM STRUCTANA WHERE GSA_AXE = "'+CodeAxe+
           '" AND GSA_ETABLISSEMENT ="' + CodeEtab + '"';
    ExecuteSQL(Req);
    TobMaj := TOB.create('Creat Ventil',NIL,-1);
    Cpt := 0;
    // Création des élements du code section
    For wi:=0 to FAffiche.Items.Count-1 do
    begin
    Inc(Cpt);
    CreatTobStruc(TObMaj, wi, Cpt, False);
    End;
    // Création des élements du libellé section
    Cpt := 0;
    For wi:=0 to FAfficheLib.Items.Count-1 do
    begin
    Inc(Cpt);
    CreatTobStruc(TObMaj, wi, Cpt, True);
    End;
    TOBMaj.InsertOrUpdateDB(False) ;
    TobMaj.free; TObMaj := NIL;
END;

procedure TOF_GCVENTILANA.CreatTobStruc (TObMaj : TOB; NumItem, Cpt : integer; TraiteLib : Boolean);
Var TobDetMaj : TOB;
    C : TlisteCol;
BEGIN
if TraiteLib then C:=TListeCol(FAfficheLib.Items.Objects[NumItem])
             else C:=TListeCol(FAffiche.Items.Objects[NumItem]) ;

TobDetMaj := Tob.create('STRUCTANA',TobMaj,-1);
TobDetMaj.PutValue('GSA_AXE',CodeAxe);
TobDetMaj.PutValue('GSA_ETABLISSEMENT',CodeEtab);
if TraiteLib    then TobDetMaj.PutValue('GSA_TYPESTRUCRANA','LSE')
                else TobDetMaj.PutValue('GSA_TYPESTRUCRANA','SEC');
TobDetMaj.PutValue('GSA_RANG',Cpt);
TobDetMaj.PutValue('GSA_LIBCHAMP',C.Champs);
TobDetMaj.PutValue('GSA_ISLIBELLE',C.IsLib) ;
TobDetMaj.PutValue('GSA_POSITION',C.Position);
TobDetMaj.PutValue('GSA_LONGUEUR',C.Longueur);
if C.sSection then TobDetMaj.PutValue('GSA_CREATSSSECTION','X')
              else TobDetMaj.PutValue('GSA_CREATSSSECTION','-');

END;

procedure TOF_GCVENTILANA.ChargeVentil;
Var CCAxe,CCEtab   : THValComboBox;
BEGIN
// Récupération des composants utilisés
CCAxe := THValComboBox (GetControl ('GSA_AXE') );

if (CCAxe.Text = '')  then
    PgiBoxAF('Vous devez saisir un axe analytique',TitreHalley)
else
    Begin
    AskMajVentil;
    ToutLiberer; Toutallouer;
    ClearEcran;
    ChargeLesTob(False,False);  ChargeLesTob(True,False); //Chargement pour le code + lib section
    ChargeLesListBox;
    SaisieEnabled (True);
    CodeAxe := CCAxe.value;
    CCEtab := THValComboBox (GetControl('GSA_ETABLISSEMENT'));
    CodeEtab := CCEtab.Value;
    MajPosition(True); MajPosition(False);
    End;
END;


procedure TOF_GCVENTILANA.OnDelete;
Var CCAxe, CCEtab : THValComboBox;
    Req : string;
BEGIN
// supression des éléments existants
If (PGIAskAF('Confirmer-vous la suppression de votre ventilation analytique ?','Ventilation analytique')= mrYes) then
   Begin
   CCAxe := THValComboBox(GetControl('GSA_AXE'));
   CCEtab:= THValComboBox(GetControl('GSA_ETABLISSEMENT'));
   Req := 'DELETE FROM STRUCTANA WHERE GSA_AXE = "'+CCAxe.Value+
       '" AND GSA_ETABLISSEMENT ="' + CCEtab.Value + '"';
   ExecuteSQL(Req);
   END;
END;

procedure TOF_GCVENTILANA.ToutAllouer;
BEGIN
TOBAffich:=TOB.Create('ventil Affich',Nil,-1) ;
TOBDispo:=TOB.Create('ventil Dispo',Nil,-1) ;
TOBAffichLib:=TOB.Create('ventil Affich Lib',Nil,-1) ;
TOBDispoLib:=TOB.Create('ventil Dispo Lib',Nil,-1) ;
END;

procedure TOF_GCVENTILANA.Toutliberer;
Var i : integer; 
BEGIN
TOBAffich.Free ; TOBAffich:=Nil ; TOBAffichLib.Free ; TOBAffichLib:=Nil ;
TOBDispo.Free ; TOBDispo:=Nil ; TOBDispoLib.Free ; TOBDispoLib:=Nil ;  
for i:=0 to FDispo.Items.Count-1 do TListeCol(FDispo.Items.Objects[i]).Free ;
for i:=0 to FAffiche.Items.Count-1 do TListeCol(FAffiche.Items.Objects[i]).Free ;
for i:=0 to FAfficheLib.Items.Count-1 do TListeCol(FAfficheLib.Items.Objects[i]).Free ;
Faffiche.Items.Clear ; FDispo.Items.Clear; FafficheLib.Items.Clear;
END;

procedure TOF_GCVENTILANA.ClearEcran;
Var CBox : TCheckBox;
BEGIN
  Fdispo.clear; Faffiche.clear;  FafficheLib.clear;  //FdispoLib.clear;  
  SetControlText('GSA_LONGUEUR',''); SetControlText('GSA_POSITION','');
  SetControlText('LONGUEURTOT','');
  SetControlText('GSA_LONGUEURLIB',''); SetControlText('GSA_POSITIONLIB','');
  SetControlText('LONGUEURTOTLIB','');
  CBox := TCheckBox(GetControl('GSA_CREATSSSECTION')); CBox.Checked := False;
END;

procedure TOF_GCVENTILANA.ChargeLesListBox;
Var  wi : integer;
      TobDet : Tob;
BEGIN
  for wi:=0  to TobDispo.Detail.count-1 do
    Begin  // boucle sur les éléments disponibles
    TobDet := TobDispo.Detail[wi];
    ChargeItem (TObDet, wi,  False,False);
    End;
  for wi:=0  to TobAffich.Detail.count-1 do
    Begin  // boucle sur les éléments affichés
    TobDet := TobAffich.Detail[wi];
    ChargeItem (TObDet, wi,  False,True);
    END;
  for wi:=0  to TobAffichLib.Detail.count-1 do
    Begin  // boucle sur les éléments affichés
    TobDet := TobAffichLib.Detail[wi];
    ChargeItem (TObDet, wi,  True,True);
    END;
END;

Procedure TOF_GCVENTILANA.SaisieEnabled (bMode : Boolean);
BEGIN
SaisieEncours := bMode;
SetControlEnabled ('BRIGHT', bMode); SetControlEnabled ('BLEFT', bMode);
SetControlEnabled ('BDOWN', bMode);  SetControlEnabled ('BUP', bMode);
SetControlEnabled ('GSA_LONGUEUR', bMode);
SetControlEnabled ('BRIGHTLIB', bMode); SetControlEnabled ('BLEFTLIB', bMode);
SetControlEnabled ('BDOWNLIB', bMode);  SetControlEnabled ('BUPLIB', bMode);
SetControlEnabled ('GSA_LONGUEURLIB', bMode);
END;

procedure TOF_GCVENTILANA.ChargeItem (TObDet:TOB; NumItem:integer; TraiteLib,Affich:Boolean);
Var
    C : TlisteCol;
BEGIN
C:= TlisteCol.Create;
C.Position := TobDet.GetValue('GSA_POSITION');
C.Longueur := TobDet.GetValue('GSA_LONGUEUR');
C.Champs   := TobDet.GetValue('GSA_LIBCHAMP');
C.IsLib := TobDet.GetValue('GSA_ISLIBELLE');
C.Lib := TobDet.GetValue('_LIB');
if Affich then C.sSection := (TobDet.GetValue ('GSA_CREATSSSECTION') = 'X');
if TraiteLib then
    BEGIN
    if Affich  then FAfficheLib.Items.AddObject(C.lib, C)
                else FDispo.Items.AddObject(C.lib, C);
    END
else
    BEGIN
    if Affich  then FAffiche.Items.AddObject(C.lib, C) 
                else FDispo.Items.AddObject(C.lib, C);
    END;
END;

procedure TOF_GCVENTILANA.ChargeLesTob (Traitelib, TobAffichOnly : Boolean) ;
Var QQ : TQuery ;
    Req,wcode,wlib,st,Champs,stWhere,Typestr,TypeTT,NomTTLibre : string;
    nbaff, Max,i,Taille : integer;
    Tobdet,TobDetDispo,TobTrt : TOB;
    CCAxe, CCEtab : THValComboBox;
    Typ,NomChamp,Lib: String ;
BEGIN
if TraiteLib then BEGIN TypeStr:='LSE'; TypeTT:='GCS'; NomTTLibre:='GCCHAMPSANA'; END
             else BEGIN Typestr:='SEC'; TypeTT:='GCS'; NomTTLibre:='GCCHAMPSANA';    END;

if Not (TobAffichOnly) then
   BEGIN
  CCAxe := THValComboBox(GetControl('GSA_AXE'));
  CCEtab:= THValComboBox(GetControl('GSA_ETABLISSEMENT'));
  Req := 'SELECT * FROM STRUCTANA WHERE GSA_TYPESTRUCRANA="'+ TypeStr +
         '" AND GSA_AXE = "'+CCAxe.Value+
         '" AND GSA_ETABLISSEMENT ="' + CCEtab.Value + '"';
  QQ:=OpenSQL(Req,True);
  If Not QQ.EOF then
    begin
    if Traitelib then
      BEGIN
      TOBAffichLib.LoadDetailDB('STRUCTANA','','',QQ,False) ;
      TOBAffichLib.Detail[0].AddChampSup('_LIB',True) ;
      TOBAffichLib.AddChampSup('_ISLIBELLE',True) ;
      nbaff := TOBAffichLib.Detail.count;
      END
    else
      BEGIN
      TOBAffich.LoadDetailDB('STRUCTRANA','','',QQ,False) ;
      TOBAffich.Detail[0].AddChampSup('_LIB',True) ;
      TOBAffich.AddChampSup('_ISLIBELLE',True) ;
      nbaff := TOBAffich.Detail.count;
      END;
    end
  else
    nbaff := 0;
  Ferme(QQ) ;
  END
else nbaff := 0;

// Recup des donnees de la tablette
if not (ctxMode in V_PGI.PGIContexte) then stWhere := FabriqueWhereItemDispo
else StWhere:= 'AND CO_LIBRE LIKE "%MODE%"' ;

Req := 'SELECT CO_CODE,CO_LIBELLE,CO_LIBRE FROM COMMUN WHERE CO_TYPE="'+TypeTT+'" ' + stWhere + ' Order by CO_CODE';
QQ:=OpenSQL(Req,True) ;
If QQ.EOF Then exit;
TobTrt := Tob.create ('liste champs', Nil,-1);
TobTrt.LoadDetailDB('','','',QQ,False);
Ferme(QQ) ;

For i := 0 To TobTrt.detail.count-1 do
   Begin
   wLib:='' ;
   wcode:= TobTrt.detail[i].GetValue('CO_CODE');
   TobTrt.detail[i].AddChampSup('_LIBCHAMP',False) ;
   Lib:=TobTrt.detail[i].GetValue('CO_LIBELLE') ;
   ReadTokenSt(Lib) ;
   if ReadTokenSt(Lib)='&#@' then wlib:=TraiteChampLibre(Lib) ;
   NomChamp:=TobTrt.detail[i].GetValue('CO_LIBRE') ;
   TobTrt.detail[i].PutValue('_LIBCHAMP',ReadTokenSt(NomChamp)) ;
   TobTrt.detail[i].AddChampSup('_TYPECHAMP',False) ;
   TobTrt.detail[i].AddChampSup('_ISLIBELLE',False) ;
   TobTrt.Detail[i].PutValue('_ISLIBELLE','-') ;
   QQ:=OpenSQL('SELECT DH_TYPECHAMP, DH_LIBELLE FROM DECHAMPS WHERE DH_NOMCHAMP="'+TobTrt.detail[i].GetValue('_LIBCHAMP')+'"',True) ;
   if not QQ.EOF then
     begin
     TobTrt.detail[i].PutValue('_TYPECHAMP',QQ.FindField('DH_TYPECHAMP').AsString) ;
     if wlib='' then wlib:=QQ.FindField('DH_LIBELLE').AsString ;
     end ;
   Ferme (QQ) ;
   Typ:=TobTrt.detail[i].GetValue('_TYPECHAMP') ;
   TobTrt.detail[i].AddChampSup('_TAILLE',False) ;
    if (copy(Typ,1,7)='VARCHAR') then Taille:=StrToInt(copy(Typ,9,length(Typ)-9))
    else if (Typ='DOUBLE') or (Typ='RATE') then Taille:=12
    else if (Typ='INTEGER') or (Typ='SMALLINT') then Taille:=8
    else if Typ='DATE' then Taille:=10
    else if Typ='COMBO' then Taille:=3 ;
    if ReadTokenSt(NomChamp)='X' then
      begin
      wLib := AdapteLibelleItems (wCode,wlib);
      Wlib:='Lib. '+Wlib ;
      Taille:=35 ;
      TobTrt.Detail[i].PutValue('_ISLIBELLE','X') ;
      end ;
   TobTrt.detail[i].PutValue('_TAILLE',Taille) ;
   if nbaff <> 0 then  // Recherche si l'element est dans la tob Affiche
      Begin
      if Traitelib then TobDet := TobAffichlib.FindFirst(['GSA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False)
                   else TobDet := TobAffich.FindFirst(['GSA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False);
      End
   else
      TobDet := NIL;
     if TobDet<>nil then begin if TobDet.GetValue('GSA_ISLIBELLE')='X' then TobDet.PutValue('_LIB','Lib. '+wlib) else TobDet.PutValue('_LIB',wlib)  end;
      if (wLib <> '.') And (wLib <> '-') then  // sinon stat non utilisée
         begin
         if traiteLib then TobDetDispo := Tob.create('STRUCTANA',TobDispoLib,-1)  
                      else TobDetDispo := Tob.create('STRUCTANA',TobDispo,-1);
         TobDetDispo.PutValue('GSA_LIBCHAMP',  TobTrt.detail[i].GetValue('_LIBCHAMP'));
         TobDetDispo.AddChampSup('_LIB',False) ;
         TobDetDispo.PutValue('_LIB', wlib);
         TobDetDispo.PutValue('GSA_LONGUEUR',TobTrt.detail[i].GetValue('_TAILLE'));
         TobDetDispo.PutValue('GSA_ISLIBELLE',TobTrt.Detail[i].GetValue('_ISLIBELLE'));
         end;
   End;
TobTrt.Free;
bModif := False;
END;

 // AC: Retourne le libelle du champ de type combo
function TOF_GCVENTILANA.TraiteChampLibre(Libelle:String):String ;
var Tablette,Code,St: String ;
begin
Result:='' ;
Tablette:=ReadTokenSt(Libelle) ;
Code:=ReadTokenSt(Libelle) ;
st := RechDom (Tablette,Code,False);
Result:=st ;
end ;

procedure TOF_GCVENTILANA.FAfficheClick(Sender: TObject);
Var C : TListeCol ;
    bModifsav : Boolean;
BEGIN
MajPosition(False);
if FAffiche.ItemIndex >= 0 then
   Begin
   bModifSav := bModif;
   C:=TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]);
   SetControlText('GSA_POSITION',IntToStr(C.Position));
   SetControlText('GSA_LONGUEUR',IntToStr (C.Longueur));
   SetControlProperty('GSA_CREATSSSECTION','Checked',C.sSection);
   bModif := bModifSav;
   END;
END;

procedure TOF_GCVENTILANA.FAffichelibClick(Sender: TObject);
Var C : TListeCol ;
    bModifsav : Boolean;
BEGIN
MajPosition(True);
if FAfficheLib.ItemIndex >= 0 then
   Begin
   bModifSav := bModif;
   C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);
   SetControlText('GSA_POSITIONLIB',IntToStr(C.Position));
   SetControlText('GSA_LONGUEURLIB',IntToStr (C.Longueur));
   bModif := bModifSav;
   END;
END;

Procedure TOF_GCVENTILANA.FRecupsSection ;
Var C : TListeCol ;
    CBox : TCheckBox;
BEGIN
if FAffiche.ItemIndex < 0 then exit;
C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);
if (C <> nil) then
    BEGIN
    CBox := TCheckBox(GetControl('GSA_CREATSSSECTION'));
    C.sSection := CBox.Checked;
    END;
END;

Procedure TOF_GCVENTILANA.FRecupLng(TraiteLib : boolean);
Var C : TListeCol ;
    suff : string;
BEGIN
if Traitelib then
    BEGIN
    if FAfficheLib.ItemIndex < 0 then exit;
    C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);
    suff :='LIB';
    END
else
    BEGIN
    if FAffiche.ItemIndex < 0 then exit;
    C:=TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]);
    Suff :='';
    END;
if C <> nil then
    BEGIN
    C.Longueur := StrToInt(GetControlText('GSA_LONGUEUR'+suff));
    if Not(MajPosition(TraiteLib)) then SetControlText('GSA_LONGUEUR'+Suff,IntToStr(C.Longueur));
    SetControlText('GSA_POSITION'+Suff,IntToStr(C.Position));
    END;
END;

// Recalcul de la position de l'item affiché et contrôles
function TOF_GCVENTILANA.MajPosition (TraiteLib : boolean) : Boolean;
Var C : TListeCol ;
    i, Posit, numError, max, MaxLng,NumI,Lng : integer ;
    Champs, NomTTLibre, Suff : string;
    FTraite : TListBox;
    TobP: Tob ;
BEGIN
if Traitelib then
    BEGIN
    FTraite := FAfficheLib;
    Suff := 'LIB'; NomTTLibre := 'GCCHAMPSANA';
    MaxLng := 35;
    END
else
    BEGIN
    FTraite := FAffiche;
    Suff := ''; NomTTLibre := 'GCCHAMPSANA';
    MaxLng := 17;
    END;
Posit := 1;  numError := 0; Result := true;
For i := 0 to FTraite.Items.Count-1 do
    BEGIN
    C:=TListeCol(FTraite.Items.Objects[i]);
    C.position :=  Posit;
    if TraiteLib then TobP:=TOBAffichLib.FindFirst(['GSA_LIBCHAMP','GSA_ISLIBELLE'],[C.Champs,C.IsLib],False)
                 else TobP:=TOBAffich.FindFirst(['GSA_LIBCHAMP','GSA_ISLIBELLE'],[C.Champs,C.IsLib],False) ;
    if tobP=nil then TobP:=TOBDispo.FindFirst(['GSA_LIBCHAMP','GSA_ISLIBELLE'],[C.Champs,C.IsLib],False) ;
    if tobP<>nil then Max := AdapteLngChamps(TobP.GetValue('GSA_LONGUEUR'),C.Champs) ;
    if C.Longueur > Max then BEGIN numError := 1; C.Longueur := Max; END;
    Posit := C.Position + C.Longueur;
    END;
Dec(Posit);
if Posit > MaxLng then numError := 2;
if  numError > 0 then
    BEGIN
    Result := False;
    PGIBoxAF (TexteMsgVentil[numError],'Ventilation Analytique');
    END;
if (NumError = 2) then
   BEGIN
   if FTraite.ItemIndex>=0 then
      BEGIN
      NumI:=FTraite.ItemIndex ;
      C:=TListeCol(FTraite.Items.Objects[NumI]) ;
      Lng := C.Longueur -(Posit - MaxLng);
      if Lng >= 0 then C.Longueur :=Lng;
      Posit := MaxLng;
      END;
   END;
SetControltext('LONGUEURTOT'+suff ,IntToStr(Posit));
END;

function TOF_GCVENTILANA.AdapteLngChamps (Lng :integer; Champs : string) : integer;
BEGIN
Result := Lng;
if Champs ='AC1' then Result :=  VH_GC.CleAffaire.Co1Lng else
 if Champs ='AC2' then Result :=  VH_GC.CleAffaire.Co2Lng else
  if Champs ='AC3' then Result :=  VH_GC.CleAffaire.Co3Lng else
  Exit;
END;

function TOF_GCVENTILANA.AdapteLibelleItems (wCode,wlib : string) : string;
Var LibLabel,Codestat : string;
BEGIN

Result := wLib; LibLabel := '';
// maj direct du libellé
  if wCode ='AC1' then Result :=  VH_GC.CleAffaire.Co1Lib else
  if wCode ='AC2' then Result :=  VH_GC.CleAffaire.Co2Lib else
  if wCode ='AC3' then Result :=  VH_GC.CleAffaire.Co3Lib else
   // Récup du libellé de la stat associée
   if wCode = 'AS1' then LibLabel := 'TAFF_LIBREAFF1' else    // Stat Affaire
   if wCode = 'AS2' then LibLabel := 'TAFF_LIBREAFF2' else
   if wCode = 'AS3' then LibLabel := 'TAFF_LIBREAFF3' else
   if wCode = 'AS4' then LibLabel := 'TAFF_LIBREAFF4' else
   if wCode = 'AS5' then LibLabel := 'TAFF_LIBREAFF5' else
   if wCode = 'AS6' then LibLabel := 'TAFF_LIBREAFF6' else
   if wCode = 'AS7' then LibLabel := 'TAFF_LIBREAFF7' else
   if wCode = 'AS8' then LibLabel := 'TAFF_LIBREAFF8' else
   if wCode = 'AS9' then LibLabel := 'TAFF_LIBREAFF9' else
   if wCode = 'ASA' then LibLabel := 'TAFF_LIBREAFFA' else
   if wCode = 'AR1' then LibLabel := 'TAFF_RESSOURCE1' else
   if wCode = 'AR2' then LibLabel := 'TAFF_RESSOURCE2' else
   if wCode = 'AR3' then LibLabel := 'TAFF_RESSOURCE3' else

   if wCode = 'PS1' then LibLabel := 'TGA_LIBREART1' else // Stat Article
   if wCode = 'PS2' then LibLabel := 'TGA_LIBREART2' else
   if wCode = 'PS3' then LibLabel := 'TGA_LIBREART3' else
   if wCode = 'PS4' then LibLabel := 'TGA_LIBREART4' else
   if wCode = 'PS5' then LibLabel := 'TGA_LIBREART5' else
   if wCode = 'PS6' then LibLabel := 'TGA_LIBREART6' else
   if wCode = 'PS7' then LibLabel := 'TGA_LIBREART7' else
   if wCode = 'PS8' then LibLabel := 'TGA_LIBREART8' else
   if wCode = 'PS9' then LibLabel := 'TGA_LIBREART9' else
   if wCode = 'PSA' then LibLabel := 'TGA_LIBREARTA' else

   if wCode = 'TS1' then LibLabel := 'TYTC_TABLELIBRETIERS1' else // Stat clients
   if wCode = 'TS2' then LibLabel := 'TYTC_TABLELIBRETIERS2' else
   if wCode = 'TS3' then LibLabel := 'TYTC_TABLELIBRETIERS3' else
   if wCode = 'TS4' then LibLabel := 'TYTC_TABLELIBRETIERS4' else
   if wCode = 'TS5' then LibLabel := 'TYTC_TABLELIBRETIERS5' else
   if wCode = 'TS6' then LibLabel := 'TYTC_TABLELIBRETIERS6' else
   if wCode = 'TS7' then LibLabel := 'TYTC_TABLELIBRETIERS7' else
   if wCode = 'TS8' then LibLabel := 'TYTC_TABLELIBRETIERS8' else
   if wCode = 'TS9' then LibLabel := 'TYTC_TABLELIBRETIERS9' else
   if wCode = 'TSA' then LibLabel := 'TYTC_TABLELIBRETIERSA' else

   if wCode = 'TR1' then LibLabel := 'TYTC_RESSOURCE1' else // Stat ressource / tiers
   if wCode = 'TR2' then LibLabel := 'TYTC_RESSOURCE2' else
   if wCode = 'TR3' then LibLabel := 'TYTC_RESSOURCE3' else
   if wCode = 'F1L' then Result :=RechDom('GCLIBFAMILLE','LF1',false) else
   if wCode = 'F2L' then Result :=RechDom('GCLIBFAMILLE','LF2',false) else
   if wCode = 'F3L' then Result :=RechDom('GCLIBFAMILLE','LF3',false)else
  Exit;

if LibLabel <> '' then // lecture du libellé de la stat ...
   BEGIN
   CodeStat := CodeTableLibre_2(LibLabel);
   If (CodeStat <> '') then
       Result :=RechDom ('GCZONELIBRE',CodeStat,False);
   END;  
END;

function TOF_GCVENTILANA.FabriqueWhereItemDispo : string;
BEGIN
Result := '';
if Not(VH_GC.CleAffaire.Co2Visible) then Result := ' AND (CO_CODE<> "AC2" AND CO_CODE<> "AC3") ';
if Not(VH_GC.CleAffaire.Co3Visible) then Result := ' AND (CO_CODE<> "AC3") ';
END;

// Gestion des déplacement
function TOF_GCVENTILANA.ctrlValidite (TraiteLib : boolean)  : boolean;
Var C,C1 : TListeCol ;
    wi : integer;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
             else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;
result := true;
C:=TListeCol(FTraiteDispo.Items.Objects[FTraiteDispo.ItemIndex]) ;
For wi:=0 to FTraiteAffiche.Items.Count-1 do
begin
  C1:=TListeCol(FTraiteAffiche.Items.Objects[wi]) ;
  if (C.Champs = C1.Champs) then
  Begin
    Result := false;
    exit;
  End;
End;
END;

procedure TOF_GCVENTILANA.BRight(TraiteLib : boolean) ;
var oldi : integer ;
    C : TListeCol ;
    ok : boolean;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
             else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;

if FTraiteDispo.ItemIndex>=0 then
   BEGIN
   OldI:=FTraiteDispo.ItemIndex ;
   C:=TListeCol(FTraiteDispo.Items.Objects[OldI]) ;
   ok := CtrlValidite(TraiteLib);
   if (ok) then
   Begin
   FTraiteAffiche.Items.AddObject(FTraiteDispo.Items[OldI],C) ;
   if OldI<FTraitedispo.Items.Count then FTraiteDispo.itemIndex:=OldI else FTraiteDispo.itemIndex:=OldI-1 ;
   FTraiteAffiche.ItemIndex:=FTraiteAffiche.Items.Count-1 ;
   if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
   End;
   END ;
end;

procedure TOF_GCVENTILANA.BLeft(TraiteLib : boolean) ;
var oldi : integer ;
    C : TListeCol ;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
             else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;

if FTraiteAffiche.ItemIndex>=0 then
  BEGIN
  OldI:=FTraiteAffiche.ItemIndex ;
  C:=TListeCol(FTraiteAffiche.Items.Objects[OldI]) ;
  FTraiteAffiche.Items.Delete(OldI) ;
  if OldI<FTraiteAffiche.Items.Count then FTraiteAffiche.itemIndex:=OldI else FTraiteAffiche.itemIndex:=OldI-1 ;
  if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
  END ;
END;

procedure TOF_GCVENTILANA.BUp(TraiteLib : boolean) ;
var oldi : integer ;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
             else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;

OldI:=FTraiteAffiche.ItemIndex ;
if OldI>0 then
   BEGIN
   if (FTraiteAffiche.Items[OldI-1]='IE_CHRONO') then Exit ;
   FTraiteAffiche.Items.Exchange(OldI,OldI-1) ;
   FTraiteAffiche.ItemIndex:=OldI-1 ;
   END ;
if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
end;

procedure TOF_GCVENTILANA.BDown(TraiteLib : boolean) ;
var oldi : integer ;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
             else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;

if (FAffiche.Items[FAffiche.ItemIndex]='IE_CHRONO') then Exit ;
OldI:=FTraiteAffiche.ItemIndex ;
if ((OldI>=0) and (oldI<FTraiteAffiche.items.Count-1)) then
   BEGIN
   FTraiteAffiche.Items.Exchange(OldI+1,OldI) ;
   FTraiteAffiche.ItemIndex:=OldI+1 ;
   END ;
if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
end;

procedure TOF_GCVENTILANA.MoveItem(Sens : String);
begin
if Not(SaisieEncours) then Exit;
bModif := True;
Uppercase (Sens);
if Sens = 'UP'   then BUp(False) else
 if Sens = 'DOWN' then BDown(False) else
  if Sens = 'LEFT' then BLeft(False) else
   if Sens = 'RIGHT' then BRight(False) else
    if Sens ='LNG' then FRecupLng(False) else
     if Sens = 'SSECTION' then FRecupsSection else
      if Sens = 'UPLIB'   then BUp(True) else
       if Sens = 'DOWNLIB' then BDown(True) else
        if Sens = 'LEFTLIB' then BLeft(True) else
         if Sens = 'RIGHTLIB' then BRight(True) else
          if Sens ='LNGLIB' then FRecupLng(True) ;
END;

procedure AGLMoveItemVentil( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFVierge) then Latof:=TFVierge(F).Latof;
if (Latof is TOF_GCVENTILANA) then
    TOF_GCVENTILANA(LaTof).MoveItem(Parms[1]);
end;

procedure AGLChargeVentil( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFVierge) then Latof:=TFVierge(F).Latof;
if (Latof is TOF_GCVENTILANA) then
    TOF_GCVENTILANA(LaTof).ChargeVentil;
end;

Initialization
registerclasses([TOF_GCVENTILANA]);
RegisterAglProc( 'MoveItemVentil',True,1,AGLMoveItemVentil);
RegisterAglProc( 'ChargeVentil',True,0,AGLChargeVentil);
end.
