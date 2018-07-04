{***********UNITE*************************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 23/01/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : REGROUPEAFFAIRE ()
Mots clefs ... : TOF;REGROUPEAFFAIRE
*****************************************************************}
Unit UtofAfRegroupeAffaire ;

Interface

Uses Forms,StdCtrls, Controls, Classes, sysutils, ComCtrls, HCtrls, HEnt1, UTOF, UTob
     ,UtofAfBaseCodeAffaire,AffaireRegroupeUtil,EntGC,SaisUtil,Vierge ,Windows
     ,AGLInitGC,LookUp,AffaireUtil,HTB97,FactUtil,AglInit,utofaffaire_mul
{$IFDEF BTP}
	 ,CalcOleGenericBTP
{$ENDIF}

{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main
{$ENDIF}
      ;
Type
  TOF_REGROUPEAFFAIRE = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  private
    IsAffaireRef, bOnLoad, IsAffaireRef_Old, IsSSAffaire, IsSSAffaire_Old, IsAffaire, IsAffaire_Old: Boolean;
    Affaire, Tiers, AffaireRef, AffaireRef_Old, stArg, MsgError : String;
    RadioBtAffRef, RadioBtSSAff, RadioBtAff : TRadioButton;
    TVAffaire : TTreeView;
    TOBAffaire, TobAffaire_Old,TobValide : TOB;
    // Gestion du grid ...
    GS : THGRID ;
    nbCol, ColAff0,ColAff1,ColAff2,ColAff3,ColAff,ColTiers : integer;
    LesCol : string;
    Action : TActionFiche;
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);

    Procedure InitRow (R : integer) ;
    procedure GereNewLigne  ;
    procedure NewLigne ;
    Function  EstRempli( Lig : integer) : Boolean ;

    procedure GSElipsisClick(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // Evènements bt liés au grid
    procedure BNEWLINEClick(Sender: TObject);
    procedure BDELLINEClick(Sender: TObject);
    procedure InsertLigne (ARow : Integer) ;
    procedure SupprimeLigne (ARow : Integer) ;
    Function  LigVide(Row : integer) : Boolean ;
    // Evenements des bt radio
    Procedure RadioBtAff_Click (Sender : TObject);
    Procedure AlimTreeView;
    Procedure PositionneTypeAff;
    Function  CtrlPositionneTypeAff : Boolean;
    Procedure ChargeTobAffaire;
    Procedure AlimGS;
    Function CtrlAvenirTobDet (Tobdet : Tob; Deviens : string) : Boolean;
    Function ValideLaSaisie: Boolean;
    Function GridToTobAffaire: Boolean;
    Function TraiteSSAffToIndependante (bTraite : Boolean): Boolean;
  end ;
 Function AFLanceFiche_RegroupAffaire(Argument:string):variant;

Implementation
/////////////// Evenements de la TOF ///////////////////////////////////////////
procedure TOF_REGROUPEAFFAIRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGROUPEAFFAIRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REGROUPEAFFAIRE.OnUpdate ;
begin
  Inherited ;
if Not(GridToTobAffaire) then
   begin
   LastError := 2; LastErrorMsg := MsgError; Exit;
   end;
TobValide := TOB.Create ('Liste tob a valider',nil,-1);
if Not(ValideLaSaisie) then
   begin
   LastError := 1; LastErrorMsg := MsgError;
   stArg := '';
   TobValide.Free;
   end
else
   begin
   TobValide.SetAllModifie(True); // TobValide.updateDB;
   if TobValide.Detail.Count > 0 then
      begin
      if TheTob <> Nil then TheTob.Free;
      TheTob:=TobValide;  // passage par theTob pour maj depuis update de l'affaire pour que tout soit enregistré dans la même transaction
      end
   else
      TobValide.Free;
   end;
end ;

procedure TOF_REGROUPEAFFAIRE.OnClose ;
begin
  Inherited ;
if TobAffaire <> Nil then TobAffaire.free;
if TobAffaire_Old <> Nil then TobAffaire_Old.free;
TFVierge(Ecran).Retour := stArg;
end;

procedure TOF_REGROUPEAFFAIRE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_REGROUPEAFFAIRE.OnArgument (StArgument : String ) ;
Var
   Critere, Champ, valeur,st : String;
   x : integer;
   BNEWLINE, BDELLINE : TToolBarButton97;
begin
Action := taModif; stArg := '';
bOnLoad := True;
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'AFF_AFFAIRE'       then Affaire := Valeur else
        if Champ = 'AFF_TIERS'         then Tiers := Valeur else
        if Champ = 'AFF_ISAFFAIREREF'  then IsAffaireRef := (Valeur='X');
        if Champ = 'AFF_AFFAIREREF'    then AffaireRef := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
if Not(IsAffaireRef) And (AffaireRef = Affaire) then AffaireRef := '';

// récup des boutons pour évènements
RadioBtAffRef := TRadioButton(Getcontrol('AFFAIREREF'));
RadioBtSSAff  := TRadioButton(Getcontrol('SSAFFAIRE'));
RadioBtAff  := TRadioButton(Getcontrol('AFFAIRE'));
RadioBtAffRef.OnClick := RadioBtAff_Click;
RadioBtSSAff.OnClick := RadioBtAff_Click;
RadioBtAff.OnClick := RadioBtAff_Click;
TVAffaire := TTreeView(GetControl('TVAFFAIRE'));

// Gestion du Grid
GS:=THGRID(GetControl('GS'));
GS.OnElipsisClick:=GSElipsisClick  ;
GS.OnDblClick:=GSDblClick ;
GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnRowEnter:=GSRowEnter ;
GS.OnRowExit:=GSRowExit ;
nbCol := 4; LesCol := 'FIXED;AFF_AFFAIRE1;';
if (VH_GC.CleAffaire.NbPartie>1) then
   BEGIN Inc(NbCol); LesCol:=LesCol + 'AFF_AFFAIRE2;' END;
if (VH_GC.CleAffaire.NbPartie>2) then
   BEGIN Inc(NbCol); LesCol:=LesCol + 'AFF_AFFAIRE3;' END;
LesCol := LesCol + 'AFF_TIERS;AFF_AFFAIRE0;';
GS.ColCount:=NbCol;

St:=LesCol;
for x:=0 to GS.ColCount-1 do
   BEGIN
   if x>2 then  GS.ColWidths[x]:=100;
   Champ:=ReadTokenSt(St) ;
   if Champ='AFF_AFFAIRE' then ColAff := x
   else if Champ='AFF_AFFAIRE0' then ColAff0 := x
   else if Champ='AFF_AFFAIRE1' then ColAff1 := x
   else if Champ='AFF_AFFAIRE2' then ColAff2 := x
   else if Champ='AFF_AFFAIRE3' then ColAff3 := x
   else if Champ='AFF_TIERS'    then ColTiers:= x   ;
   END ;
// libellé des colonnes
GS.Cells[ColTiers,0]:= 'Client';
GS.Cells[ColAff1,0]:= VH_GC.CleAffaire.Co1Lib;
if ColAff2 <> 0 then GS.Cells[ColAff2,0]:= VH_GC.CleAffaire.Co2Lib;
if ColAff3 <> 0 then GS.Cells[ColAff3,0]:= VH_GC.CleAffaire.Co3Lib;
GS.ColWidths[0]:=15;  // Fixed col
GS.ColWidths[ColAff0]:=0;  // affaire 0
AffecteGrid(GS,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
GS.ColWidths[0]:=15;  // Fixed col

// Boutons liés au grid
BNEWLINE:=TToolbarButton97(GetControl('BNEWLINE'));
BNEWLINE.OnClick:=BNEWLINECLick;
BDELLINE:=TToolbarButton97(GetControl('BDELLINE'));
BDELLINE.OnClick:=BDELLINECLick;

// positionnement initial des champs
IsSSAffaire := (Affaire <> AffaireRef) And (AffaireRef <> '');
IsAffaire := Not(IsAffaireRef) And Not(IsSSAffaire);
SetControlText('AFF_AFFAIRE', Affaire); SetControlText('AFF_AFFAIREREF', AffaireRef);
SetControlText('AFF_TIERS', Tiers);

SetControlChecked('AFFAIREREF',IsAffaireRef);
SetControlChecked('SSAFFAIRE',IsSSAffaire);
SetControlChecked('AFFAIRE',IsAffaire);

// valeurs d'entrée conservées pour test validation + contrôle
IsAffaireRef_Old := IsAffaireRef; AffaireRef_old := AffaireRef;
IsSSAffaire_Old := IsSSAffaire;  IsAffaire_old := IsAffaire;
bOnload:=False;
Inherited;
END;


procedure TOF_REGROUPEAFFAIRE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));   Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2')); Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));  Tiers:=THEdit(GetControl('AFF_TIERS'));

Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));   Aff1_:=THEdit(GetControl('AFF_AFFAIREREF1'));
Aff2_:=THEdit(GetControl('AFF_AFFAIREREF2')); Aff3_:=THEdit(GetControl('AFF_AFFAIREREF3'));
Aff4_:=THEdit(GetControl('AFF_AVENANTREF'));  Tiers_:=THEdit(GetControl('AFF_TIERSREF'));
end;
/////////////////////////// Evenements Bt Affaire / sous Affaire /////////////
Procedure TOF_REGROUPEAFFAIRE.RadioBtAff_Click (Sender : TObject);
BEGIN
PositionneTypeAff;
END;

////// Gestion du Grid ///////
procedure TOF_REGROUPEAFFAIRE.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
GereNewLigne;
if Not Cancel then
   BEGIN
   GS.ElipsisButton := False;
   if (GS.Col=ColAff0) then GS.ElipsisButton :=True else
   if (GS.Col=ColAff1) And (VH_GC.CleAffaire.Co1Type='LIS') then GS.ElipsisButton := true else
   if (GS.Col=ColAff2) And (VH_GC.CleAffaire.Co2Type='LIS') then GS.ElipsisButton := true else
   if (GS.Col=ColAff3) And (VH_GC.CleAffaire.Co3Type='LIS') then GS.ElipsisButton := true else
   if (GS.Col=ColTiers) then GS.ElipsisButton :=True;
   end ;
end;

procedure TOF_REGROUPEAFFAIRE.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
end;


procedure TOF_REGROUPEAFFAIRE.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_REGROUPEAFFAIRE.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Not EstRempli(ou) then GS.DeleteRow(ou) ;
GS.InvalidateRow(ou) ;
end;

//////////// Gestion des évenements du Grid  (Recherche sur elipsis du Grid...)
procedure TOF_REGROUPEAFFAIRE.GSElipsisClick(Sender: TObject);
var value : string ;
begin
if GS.Col=ColAff0 then LookUpPartiesAffaire( 0, GS) else
if GS.Col=ColAff1 then LookUpPartiesAffaire( 1, GS) else
if GS.Col=ColAff2 then LookUpPartiesAffaire( 2, GS) else
if GS.Col=ColAff3 then LookUpPartiesAffaire( 3, GS) else
if GS.Col=ColTiers then
   BEGIN
   value:=DispatchRecherche(Nil,2,'T_NATUREAUXI="CLI"','','');
   if Value <> '' then GS.Cells[ColTiers,GS.Row] := Value;
   END;
end;

procedure TOF_REGROUPEAFFAIRE.GSDblClick(Sender: TObject);
var Value,Stchamps,Stargument,Tiers,Aff,Aff0,Aff1,Aff2,Aff3,Aff4 : string ;
begin
Tiers:=GS.Cells[ColTiers,GS.Row];
if (GS.Col=ColAff1) or (GS.Col=ColAff1) or (GS.Col=ColAff2) or (GS.Col=ColAff3) then
   BEGIN
   Stchamps :='AFF_STATUTAFFAIRE=AFF';
   Stargument := 'NOCHANGESTATUT';
   // Stargument := Stargument + ';NOCHANGETIERS';
   if (Tiers <> '') then StChamps:=StChamps+';AFF_TIERS='+Tiers;
   // mcd 12/06/02 Value:=AGLLanceFiche('AFF','AFFAIRERECH_MUL',StChamps,'',StArgument) ;
   Value:=  AFLanceFiche_AffaireRech(StCHamps,StArgument);
   if Value <>'' then
      BEGIN
      Aff := ReadTokenSt(Value); // retourne affaire;tiers
      {$IFDEF BTP}
      BTPCodeAffaireDecoupe(Aff,Aff0,Aff1,Aff2,Aff3,Aff4, taCreat, false);
      {$ELSE}
      CodeAffaireDecoupe(Aff,Aff0,Aff1,Aff2,Aff3,Aff4, taCreat, false);
      {$ENDIF}
      GS.Cells[ColAff0,GS.Row] := Aff0; GS.Cells[ColAff1,GS.Row] := Aff1;
      GS.Cells[ColAff2,GS.Row] := Aff2; GS.Cells[ColAff3,GS.Row] := Aff3;
      GS.Cells[ColTiers,GS.Row] := Value;
      END;
   END
else GSElipsisClick(sender);
end;

procedure TOF_REGROUPEAFFAIRE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
BEGIN
if Action=taConsult then Exit ;
FocusGrid := False;  ARow:=0;
if(Screen.ActiveControl = GS) then
    BEGIN
    FocusGrid := True;
    ARow := GS.Row;
    END ;
Case Key of
    VK_F5 : if (Screen.ActiveControl=GS) then GSElipsisClick(Sender);
    VK_RETURN : Key:=VK_TAB ;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow) ;
                    END ;
                END;
    END;
end;


////////////// Gestion des nouvelles lignes ////////////////////////////////////
procedure TOF_REGROUPEAFFAIRE.GereNewLigne  ;
BEGIN
   if EstRempli(GS.RowCount-1) then NewLigne else
      if Not EstRempli(GS.RowCount-2) then GS.RowCount:=GS.RowCount-1 ;
END ;

procedure TOF_REGROUPEAFFAIRE.NewLigne ;
BEGIN
    GS.RowCount:=GS.RowCount+1 ;
    InitRow( GS.RowCount-1 ) ;
END ;

Procedure TOF_REGROUPEAFFAIRE.InitRow (R : integer) ;
var i : integer ;
begin
for i:=0 to GS.ColCount do GS.cells[i,R]:='';
end ;

Function TOF_REGROUPEAFFAIRE.EstRempli( Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
for i:=1 to GS.ColCount-1 do
  if (GS.Cells[i,Lig]<>'') then begin result:= true; break; end;
END ;

/////// GESTION Insertion, Suppression de lignes ///////////////////////////////
procedure TOF_REGROUPEAFFAIRE.BNEWLINEClick(Sender: TObject);
begin
if Action=taConsult then Exit ;
InsertLigne (GS.Row);
end;

procedure TOF_REGROUPEAFFAIRE.BDELLINEClick(Sender: TObject);
begin
if Action=taConsult then Exit ;
SupprimeLigne (GS.Row) ;
end;

procedure TOF_REGROUPEAFFAIRE.InsertLigne (ARow : Integer) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigVide (ARow) then exit;
GS.CacheEdit; GS.SynEnabled := False;
GS.InsertRow (ARow); GS.Row := ARow; GS.Col:=ColAff1;
InitRow (ARow) ;
GS.MontreEdit; GS.SynEnabled := True;
END;

procedure TOF_REGROUPEAFFAIRE.SupprimeLigne (ARow : Integer) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
GS.CacheEdit; GS.SynEnabled := False;
GS.DeleteRow (ARow);
GS.MontreEdit; GS.SynEnabled := True;
END;

Function TOF_REGROUPEAFFAIRE.LigVide(Row : integer) : Boolean ;
BEGIN
Result:=True ;
if (GS.Cells[ColAff1,Row]<>'') then result:= False ;
END ;


/////////////// CHARGEMENT Affaire / Sous affaire //////////////////////////////
Procedure TOF_REGROUPEAFFAIRE.PositionneTypeAff;
BEGIN
IsAffaireRef := RadioBtAffRef.Checked; IsAffaire := RadioBtAff.Checked; IsSSAffaire :=RadioBtSSAff.Checked;
if Not(bOnload) then
   if Not(CtrlPositionneTypeAff) then Exit;
// Images affichées
SetControlVisible ('IMAGEAFFREF',IsAffaireRef); SetControlVisible ('IMAGESSAFF',(IsSSAffaire));
SetControlVisible ('IMAGEAFF',IsAffaire);
// Grid
SetControlVisible ('PGRID',IsAffaireRef);
SetControlVisible ('BNEWLINE',IsAffaireRef); SetControlVisible ('BDELLINE',IsAffaireRef);
SetControlVisible ('PCLIENT', Not(IsAffaire)); //SetControlVisible ('TVAFFAIRE', Not(IsAffaire));

if IsAffaireRef then SetControlProperty('PENTETECLIENT','Caption', 'Liste des sous affaires associées')
                else SetControlProperty('PENTETECLIENT','Caption', 'Affaire de référence associée');
ChargeTobAffaire;

if bOnLoad then
   begin
   if TobAffaire_Old = Nil then TobAffaire_Old := TOB.Create('',Nil,-1);
   TobAffaire_Old.Dupliquer (TobAffaire,True,True);
   if IsSSAffaire then SetControlText('AFF_TIERSREF',TobAffaire.GetValue('AFF_TIERS'));
   end;
AlimTreeView;
AlimGS;
END;


Procedure TOF_REGROUPEAFFAIRE.ChargeTobAffaire;
Var //ssAffDetect : integer;
    //Q : TQuery;
    //bChargeAffRef : Boolean;
    TobEnc : TOB;
BEGIN
if (IsAffaireRef) or (IsAffaire) then // Chargement des affaires ref + sous affaires associées
   BEGIN
   if TobAffaire <> Nil then
      begin
      if TobAffaire.GetValue('AFF_AFFAIRE') = Affaire then
         begin
         if TobAffaire.Detail.Count > 0 then TobAffaire.ClearDetail;
         end
      else
         TobAffaire.Free; TobAffaire := Nil;
      end;
   if TobAffaire = Nil then
         begin
         TobAffaire := Tob.Create('AFFAIRE',Nil,-1);
         ChargeTobAffaireRef(TobAffaire,Affaire,'',True);
         end;
   if IsAffaireRef then
     //ssAffDetect := ChargeTobSSAffaire (TobAffaire, Affaire,'',True);
     ChargeTobSSAffaire (TobAffaire, Affaire,'',True);
   END
else if isSSAffaire then
   BEGIN
   // Recup affaire référence si existe + les autres sous affaires
   if affaireRef <> '' then
      BEGIN
      if TobAffaire <> Nil then TobAffaire.Free;
      TobAffaire := Tob.Create('AFFAIRE',Nil,-1);
      ChargeTobAffaireRef(TobAffaire,AffaireRef,'',True);
      //ssAffDetect :=ChargeTobSSAffaire (TobAffaire, AffaireRef,'',True);
      ChargeTobSSAffaire (TobAffaire, AffaireRef,'',True);
      END
   else
      BEGIN
      TobEnc := Tob.Create('sav affaire',Nil,-1); TobEnc.Dupliquer(TobAffaire,False,True);
      if TobAffaire <> Nil then TobAffaire.Free;
      TobAffaire := Tob.Create('AFFAIRE',Nil,-1); TobAffaire.AddChampsup('AFFAIREAFFICHE',True);// Affaire de réf ...
      TobEnc.ChangeParent(TobAffaire,-1);
      END;
   END;
END;

Procedure TOF_REGROUPEAFFAIRE.AlimGS;
BEGIN
if TobAffaire = Nil then Exit;
if TobAffaire.Detail.Count=0 then Exit;
if IsAffaire Then Exit;
TobAffaire.PutGridDetail(GS,False,False,LesCol,True);
GS.RowCount:=GS.RowCount+1 ;
END;

Procedure TOF_REGROUPEAFFAIRE.AlimTreeView;
BEGIN
if TobAffaire=Nil then Exit; //if IsAffaire Then Exit;
if TVAffaire <> Nil then
   begin
   TobAffaire.PutTreeView(TVAffaire, Nil,'"Affaire ref "|AFFAIREAFFICHE|" - client "|AFF_TIERS;AFFAIREAFFICHE|" - Client "|AFF_TIERS', 2, 2);
   TVAffaire.FullExpand;
   end;
END;

/////////////////////// Contrôles Affaires / sous Affaires //////////////////////
Function  TOF_REGROUPEAFFAIRE.CtrlPositionneTypeAff : Boolean;
BEGIN
Result := True;
END;

////////////////////////Validation /////////////////////////////////////////////
Function TOF_REGROUPEAFFAIRE.CtrlAvenirTobDet(TobDet : TOB; Deviens : string): Boolean;
var TiersRef : string;
BEGIN
// Deviens 'AFF' : affaire indépendante 'SSA' : Sous Affaire 'REF' : Affaire Ref
Result := True;
TiersRef := TobAffaire.GetValue('AFF_TIERS');

// **********************  deviens une sous affaire ****************************
if deviens = 'SSA' then
   begin
   
   // passage Affaire Ref => Sous affaire interdit (pb multi niveau ...)
   if (TobDet.Getvalue('AFF_ISAFFAIREREF')='X') then
      begin
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' était une affaire de référence. Elle ne peut être prise comme sous-affaire';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' était une affaire de référence. Elle ne peut être prise comme sous-affaire';
      {$ENDIF}
      Result:= false;
      Exit;
      end;

   // Si facturation sur code de regoupement 1,2,3 => affaire reste indépendante
   if Copy(TobDet.Getvalue('AFF_REGROUPEFACT'),1,2)= 'RE' then // sur code de regroupement
      begin
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur un code de regroupement indépendant. Elle ne peut être intégrée en sous affaire';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur un code de regroupement indépendant. Elle ne peut être intégrée en sous affaire';
      {$ENDIF}
      Result:= false; Exit;
      end;

   // Si facturation regroupée sur affaire référence + tiers différent de l'aff référence
   if (TobDet.Getvalue('AFF_REGROUPEFACT')= 'ARE') and (TiersRef <>TobDet.Getvalue('AFF_TIERS'))  then // fact. sur affaire de regroupement
      begin
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur l''affaire de référence avec un tiers différent';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur l''affaire de référence avec un tiers différent';
      {$ENDIF}
      Result:= false; Exit;
      end;
   end else
// ************************   deviens une affaire référence ********************
if deviens = 'REF' then
   begin
   // passage directe sous-affaire => affaire de référence interdit (pb multi-niveau)
   if ( TobDet.Getvalue('AFF_AFFAIREREF') <> TobDet.Getvalue('AFF_AFFAIRE') ) then
      begin
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' était une affaire sous-affaire. Elle ne peut être prise comme affaire de référence';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' était une affaire sous-affaire. Elle ne peut être prise comme affaire de référence';
      {$ENDIF}
      Result:= false; Exit;
      end;
   // si aucune sous affaire, on ne devient pas une affaire de référence
   if TobAffaire.Detail.Count = 0 then
      begin
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' ne peut pas être une affaire référence sans sous-affaires';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' ne peut pas être une affaire référence sans sous-affaires';
      {$ENDIF}
      Result:= false; Exit;
      end;
   end else
// ******************************** deviens une affaire ***********************
if deviens = 'AFF' then
   begin
   // si facturation sur l'affaire de référence (indépendance impossible)
   if TobDet.Getvalue('AFF_REGROUPEFACT')= 'ARE' then // sur affaire de référence
      begin
      result:= false;
      {$IFDEF BTP}
      msgError := 'L''affaire '+ BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur l''affaire de référence. Elle ne peut pas être indépendante';
      {$ELSE}
      msgError := 'L''affaire '+ CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' ') +
      ' est facturée sur l''affaire de référence. Elle ne peut pas être indépendante';
      {$ENDIF}
      end;
   end;
END;

Function TOF_REGROUPEAFFAIRE.ValideLaSaisie: Boolean;
Var i,j : integer;
    TobDet,TobDet_Old,{TobSSAff,}TobDet_Valide : TOB;
    bTrouve : Boolean;
BEGIN
Result := True;
stArg := '';
// on repasse en argument les valeurs à modifier sur l'affaire en cours + maj SQL des autres affaires.

//************************** Affaire de référence  ********************************
if IsAffaireRef then
   BEGIN
   // Traitement de l'affaire de référence
   if Not(IsAffaireRef_Old) then
      begin
      if Not(CtrlAvenirTobDet(TobAffaire ,'REF')) then begin Result := false; exit; end;
      stArg := 'AFF_ISAFFAIREREF:X;AFF_AFFAIREREF:'+ Affaire;
      end;
   // Traitement des sous affaires
   for i := TobAffaire.detail.count-1 downto 0 do
      begin
      TobDet := TobAffaire.detail[i];
      // recherche si cette sous affaire déja rattachée à l'affaire de référence
      bTrouve := false;

      for j:= 0 to TobAffaire_Old.detail.count-1 do
         begin
         TobDet_Old :=  TobAffaire_Old.detail[j];
         if j=0 then TobDet_Old.AddChampSup('TRAITE',True);
         if TobDet.GetValue('AFF_AFFAIRE')=TobDet_Old.GetValue('AFF_AFFAIRE') then
            begin
            TobDet_Old.PutValue('TRAITE','X');
            bTrouve:=True; break;
            end;
         end;
      // si elle n'etait pas rattachée auparavant => maj
      if Not(btrouve) then
         begin
         if Not(CtrlAvenirTobDet(TobDet ,'SSA')) then begin Result := false; exit; end;
         TobDet.PutValue('AFF_ISAFFAIREREF','-'); TobDet.PutValue('AFF_AFFAIREREF',Affaire);
         TobDet.ChangeParent(TobValide,-1);
         end;
      end;
      if Not TraiteSSAffToIndependante(True) then begin Result:=false; Exit; end;
   END

//******************************* Sous Affaire *********************************
else if IsSSAffaire then
   BEGIN
   if Not((IsSSAffaire_Old) And (AffaireRef = AffaireRef_Old)) then
      begin
      // maj de la sous affaire
      Tobdet := Nil;
      for i := TobAffaire.detail.count-1 downto 0 do
         begin
         TobDet := TobAffaire.detail[i];
         if TobDet.GetValue('AFF_AFFAIRE')=Affaire then break;
         end;
      if TobDet = Nil then
         begin TobDet := Tob.Create('AFFAIRE',TobAffaire,-1);  RemplirTobAffaire (Affaire,TobDet); end;

      if Not(CtrlAvenirTobDet(TobDet ,'SSA')) then begin Result := false; exit; end;
      stArg := 'AFF_ISAFFAIREREF:-;AFF_AFFAIREREF:'+ AffaireRef;
      // maj de l'affaire de référence
      if Not(CtrlAvenirTobDet(TobAffaire ,'REF')) then begin Result := false; exit; end;
      TobAffaire.PutValue('AFF_ISAFFAIREREF','X'); TobAffaire.PutValue('AFF_AFFAIREREF',TobAffaire.GetValue('AFF_AFFAIRE'));
      TobDet_Valide := Tob.Create('AFFAIRE',TobValide,-1);
      TobDet_Valide.Dupliquer(TobAffaire,False,True);
      end;
   END

//*************************** Affaire indépendante *****************************
else if IsAffaire then
   BEGIN
   if Not(IsAffaire_Old) then
      begin
      if Not(CtrlAvenirTobDet(TobAffaire ,'AFF')) then begin Result := false; exit; end;
      stArg := 'AFF_ISAFFAIREREF:-;AFF_AFFAIREREF:'+ Affaire;
      if IsAffaireRef_Old then
         if Not TraiteSSAffToIndependante(false) then begin Result:=false; Exit; end;
      if IsSSAffaire_Old then
         begin
         bTrouve :=False;
         for j:= 0 to TobAffaire_Old.detail.count-1 do
            begin
            TobDet_Old :=  TobAffaire_Old.detail[j];
            if TobDet_Old.GetValue('AFF_AFFAIRE')<> Affaire then bTrouve := True;
            end;
         if Not(bTrouve) then // aucune affaire diff. trouvée sur affaire de réference => deviens indépendante
            begin
            TobAffaire_Old.PutValue('AFF_ISAFFAIREREF','-');
            TobDet := Tob.Create('aff ref',TobValide,-1);
            TobDet.Dupliquer (TobAffaire_Old,False,True);
            end;
         end;
      end;
   END;
END;

Function TOF_REGROUPEAFFAIRE.GridToTobAffaire : Boolean;
Var i , j : integer;
    Aff0,Aff1,Aff2,Aff3,Aff4,Tiers,CodeAff : string;
    TobA,TobDet : TOB;
    bTrouve : Boolean;
BEGIN
Result := False;
if TobAffaire=Nil then Exit;
if IsAffaire then begin Result := True; Exit; end;
if TobAffaire.Detail.Count > 0 then TobAffaire.detail[0].AddChampSup('CHARGE',True);

if IsSSAffaire then
   begin   // Sous affaire => Recup dans lesTHEdit
      Aff0:='';Aff1:='';Aff2:='';Aff3:='';Aff4:='';Tiers:='';CodeAff:='';
      Aff0:=GetControlText('AFF_AFFAIREREF0'); if Aff0='' then Aff0 := 'A';
      Aff1:=GetControlText('AFF_AFFAIREREF1');
      if (VH_GC.CleAffaire.NbPartie>1) then Aff2:=GetControlText('AFF_AFFAIREREF2');
      if (VH_GC.CleAffaire.NbPartie>1) then Aff3:=GetControlText('AFF_AFFAIREREF3');
      Aff4:='00';
      Tiers := GetControlText('AFF_TIERSREF');
      if Aff1 = '' then begin MsgError:= 'Affaire référence non renseignée'; exit; end;

      CodeAff := CodeAffaireRegroupe(Aff0,Aff1,Aff2,Aff3,Aff4, Action,False,False,False);
      if TeststCleAffaire(CodeAff,Aff0,Aff1,Aff2,Aff3,Aff4,Tiers, false,false,false,true)<>1 then
         begin
         {$IFDEF BTP}
         MsgError:= 'Code Affaire référence invalide: ' + BTPCodeAffaireAffiche(CodeAff);
         {$ELSE}
         MsgError:= 'Code Affaire référence invalide: ' + CodeAffaireAffiche(CodeAff);
         {$ENDIF}
         exit;
         end;
      if CodeAff = Affaire then
         begin
         {$IFDEF BTP}
         MsgError:= 'Code Affaire référence invalide: '+BTPCodeAffaireAffiche(CodeAff)+'. La sous affaire = l''affaire de référence';
         {$ELSE}
         MsgError:= 'Code Affaire référence invalide: '+CodeAffaireAffiche(CodeAff)+'. La sous affaire = l''affaire de référence';
         {$ENDIF}
         exit;
         end;

      // codeAff présent dans la TobAffaire
      btrouve := False;
      if Codeaff = TobAffaire.GetValue('AFF_AFFAIRE') then btrouve:=true;
      if Not (bTrouve) then
         begin
         TobAffaire.InitValeurs;
         RemplirTOBAffaire (CodeAff,TobAffaire);
         end;
      AffaireRef := codeAff;
   end
else
   begin   // Affaire référence => récup dans le grid
   for i := 1 to GS.Rowcount do
      begin
      Aff0:='';Aff1:='';Aff2:='';Aff3:='';Aff4:='';Tiers:='';CodeAff:='';
      Aff0:=GS.Cells[ColAff0,i]; if Aff0='' then Aff0 := 'A';
      Aff1:=GS.Cells[ColAff1,i];
      if (VH_GC.CleAffaire.NbPartie>1) then Aff2:=GS.Cells[ColAff2,i];
      if (VH_GC.CleAffaire.NbPartie>1) then Aff3:=GS.Cells[ColAff3,i];
      Aff4:='00';
      Tiers := GS.Cells[ColTiers,i];
      if Aff1 = '' then continue;
      CodeAff := CodeAffaireRegroupe(Aff0,Aff1,Aff2,Aff3,Aff4, Action,False,False,False);
      if TeststCleAffaire(CodeAff,Aff0,Aff1,Aff2,Aff3,Aff4,Tiers, false,false,false,True)<>1 then
         begin
         {$IFDEF BTP}
         MsgError:= 'Code Affaire invalide: ' + BTPCodeAffaireAffiche(CodeAff);
         {$ELSE}
         MsgError:= 'Code Affaire invalide: ' + CodeAffaireAffiche(CodeAff);
         {$ENDIF}
         exit;
         end;
      if CodeAff = Affaire then
         begin
         {$IFDEF BTP}
         MsgError:= 'Code Affaire invalide: '+BTPCodeAffaireAffiche(CodeAff)+'. La sous affaire = l''affaire de référence';
         {$ELSE}
         MsgError:= 'Code Affaire invalide: '+CodeAffaireAffiche(CodeAff)+'. La sous affaire = l''affaire de référence';
         {$ENDIF}
         exit;
         end;

      // codeAff présent dans la TobAffaire
      btrouve := False;
      for j := 0 to TobAffaire.detail.Count-1 do
         begin
         if Codeaff = TobAffaire.detail[j].GetValue('AFF_AFFAIRE') then
            begin btrouve:=true; TobAffaire.detail[j].PutValue('CHARGE','X'); break; end;
         end;
      if Not (bTrouve) then
         begin
         TobA := TOB.Create('AFFAIRE',Nil,-1);
         TobA.AddChampSup('CHARGE',False); TobA.PutValue('CHARGE','X');
         if RemplirTOBAffaire (CodeAff,TobA) then TobA.ChangeParent(TobAffaire,-1);
         end;
      // Suppression de celles non utilisées
      for j := TobAffaire.detail.count-1 downto 0 do
         begin
         TobDet := TobAffaire.detail[j];
         if TobDet.GetValue('CHARGE') <>'X' then TobDet.Free;
         end;
      end;
   end;
Result := true;
END;

Function TOF_REGROUPEAFFAIRE.TraiteSSAffToIndependante (bTraite : Boolean): Boolean;
Var j : integer;
    TobDet_Old: TOB;
BEGIN
Result := true;
// Celles qui étaient en sous affaires et qui deviennent indépendantes (freedom ...)
for j := TobAffaire_Old.detail.count-1 downto 0 do
   begin
   TobDet_Old :=  TobAffaire_Old.detail[j];
   if bTraite then if TobDet_Old.GetValue('TRAITE') = 'X' then Continue;
   if Not(CtrlAvenirTobDet(TobDet_Old ,'AFF')) then begin Result := false; exit; end;
   TobDet_Old.PutValue('AFF_ISAFFAIREREF','-'); TobDet_old.PutValue('AFF_AFFAIREREF',TobDet_old.GetValue('AFF_AFFAIRE'));
   TobDet_old.ChangeParent(TobValide,-1);
   end;
END;

Function AFLanceFiche_RegroupAffaire(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFREGROUPEAFFAIRE','','',Argument);
end;

Initialization
registerclasses ( [ TOF_REGROUPEAFFAIRE ] ) ;
end.
