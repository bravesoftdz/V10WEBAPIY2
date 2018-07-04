{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 15/05/2006
Modifié le ... : 22/05/2006
Description .. : Source TOF de la FICHE : UTOFCIBLAGE ()
Suite ........ : Liste des Ciblages
Suite ........ : Ecran de création des ciblages
Suite ........ : Choix des cibles
Mots clefs ... : TOF;UTOFCIBLAGE
*****************************************************************}
Unit UTOFCIBLAGE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     utilGc,    //changelibre2
     ed_tools,  //InitMoveProgressForm
     HMsgBox,
     utofcomm,
     KpmgUtil,
     UTOF,
     HTB97,
     {$ifdef AFFAIRE}
     UTOFAFTRADUCCHAMPLIBRE,
     {$endif}
     uTofRTTOPCONTACT_TOF,
     utofRTTOPTIERSCON_TOF,
     UtofRTTOPSUSCON_TOF,
     UTofRTTOPCIBLE,
     UTofRTTOPTIERS_TOF,
     uTofRTCIBLESUS_TOF
      ;

Type
{$ifdef AFFAIRE}
                //PL le 25/05/07 pour gérer les tablettes en correspondance
     TOF_UTOFCIBLAGE = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_UTOFCIBLAGE = Class (tTOFComm)
{$endif}
   private

   public
    LaListe : THGrid;
    Bselect : TToolBarButton97;
    Binsert : TToolBarButton97;
    Rb1                 : TRadioButton;
    Rb2                 : TRadioButton;
    Rb3                 : TRadioButton;
    Bbloque             : Boolean;
    BAvert              : Boolean;


    procedure BselectOnClick ( Sender : TObject );
    procedure LaListeOnFlipSelection ( Sender : TObject );
    procedure LaListeOnDblClick ( Sender : TObject );
    procedure ZoomOnClick(Sender: TObject);
    procedure BinsertOnClick ( Sender : TObject );
    procedure BsupprimerOnClick (Sender : Tobject);
    procedure RbOnClick (Sender : Tobject);


    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_utofciblage(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_utofciblage(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_UTOFCIBLAGE.OnArgument (S : String ) ;
var
   //Q : TQuery;
   i : integer;
   //LaGrille : THGrid;
   Cbval  : THCheckbox;
   Bsuppr, BZoom : TToolbarButton97;
   F                    : TFMul;
   StrMsg               : String;

begin
  Inherited ;
//FQ 10484 TJA 26/02/2008
  F                     := TFMul(ecran);
  StrMsg                := JaiLeDroitCiblage(Bbloque, BAvert);

  Rb1                   := TRadioButton(GetControl('RB1'));
  Rb2                   := TRadioButton(GetControl('RB2'));
  Rb3                   := TRadioButton(GetControl('RB3'));


  if Assigned(Rb1) And Assigned(Rb2) and Assigned(Rb3) then
  begin
    Rb1.OnClick         := RbOnClick;
    Rb2.OnClick         := RbOnClick;
    Rb3.OnClick         := RbOnClick;
    SetControlText('RCB_MODECIBLAGE', '001');
  end;

  InitRadioBtModeCiblage(F, Bbloque);

  if S = 'SUPPRESSION' then
  begin
    TFMul(Ecran).Caption := 'Suppression ciblage';
    UpdateCaption(Tfmul(Ecran));
    SetControlVisible('BOUVRIR', False);
    SetControlVisible('BINSERT', False);
    SetControlVisible('BSELECT', False);
    SetControlVisible('BZOOM', true);
    
    BZoom := TToolBarButton97(getcontrol('BZOOM'));
    BZoom.OnClick := ZoomOnClick;

    Bsuppr  := TToolbarButton97(GetControl('BSUPPRIMER'));
    if Assigned(Bsuppr) then
    begin
      Bsuppr.Visible  := True;
      Bsuppr.OnClick  := BsupprimerOnClick;
    end;
    SetControlVisible('BSELECTALL', True);
    //LaGrille  := THGrid(GetControl('FLISTE'));
    //LaGrille.MultiSelect := True;
    Cbval   := THCheckbox(GetControl('RCB_VALIDATION'));
    Cbval.State := cbGrayed;
    fTableName := 'CIBLAGE' ;
    fMulDeTraitement := True;
    SetMultiSelection(true);


  end
  else
  begin
    //SetControlVisible('BOUVRIR', False);

    Bselect := TToolBarButton97(getcontrol('BSELECT'));
    Bselect.OnClick := BselectOnClick;
    // PL le 17/04/06 : bouton ouvrir actif = dblclick
    Bselect := TToolBarButton97(getcontrol('BOUVRIR'));
    Bselect.OnClick := LaListeOnDblClick;


    Binsert := TToolBarButton97(GetControl('BINSERT'));
    Binsert.OnClick := BinsertOnClick;

    LaListe := THGrid(GetControl('FLISTE'));
    LaListe.OnDblClick := LaListeOnDblClick;
    SetControlVisible ('BINSERT', True);

    //Q:=OpenSQL('SELECT MN_LIBELLE FROM MENU WHERE MN_TAG="92261"', True);
    //if not Q.EOF then
    //begin
      //TFMul(Ecran).Caption := Q.Fields[0].AsString;
    // PL le 13/04/07 : on peut mettre le titre en dur : pas de requête et il faut ajouter
    // le mot "ciblage" dedans
      TFMul(Ecran).Caption := 'Création / sélection ciblage';
      UpdateCaption(TFMul(ecran));

      if StrMsg <> '' then
        PGIBox(StrMsg, TFMul(Ecran).Caption);
    //end;

    //Ferme(Q);
  end;
  // maj des champs libres
   for i := 1 to 5 do
   begin
    ChangeLibre2('TRCB_RCBTABLELIBRE'+intToStr(i), Tform(Ecran));
    ChangeLibre2('TRCB_RCBTEXTELIBRE'+intToStr(i), Tform(Ecran));
    ChangeLibre2('TRCB_RCBDATELIBRE'+intToStr(i), Tform(Ecran));
   end;

end ;

procedure TOF_UTOFCIBLAGE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCIBLAGE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCIBLAGE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCIBLAGE.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_UTOFCIBLAGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCIBLAGE.OnDisplay () ;
begin
  Inherited ;
//   LaListeOnFlipSelection(nil);

end ;

procedure TOF_UTOFCIBLAGE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/05/2006
Modifié le ... :   /  /    
Description .. : Appel de la fiche ciblage pour Raison Sociale
Mots clefs ... : 
*****************************************************************}
procedure TOF_UTOFCIBLAGE.BselectOnClick(Sender: TObject);
var
  LeMode                : String;
  LeType                : String;
  LeCode                : String;
//  StrMsg                : String;
//  bBloque               : boolean;

begin
  //verif validation ciblage
  if GetField('RCB_VALIDATION')='X' then
  begin
    PGIBox('Ciblage déjà validé');
    exit;
  end;

  if VarIsNull(GetField('RCB_CODECIBLAGE')) then
    exit;

  LeMode                := GetField('RCB_MODECIBLAGE');
  LeType                := GetField('RCB_TYPECIBLAGE');
  LeCode                := GetField('RCB_CODECIBLAGE');

// FQ 10484 TJA 26/02/2008
  //FQ 10484
  if ((LeMode = '003') or (LeMode = '002')) and Bbloque then
      exit;

  //si ciblage sur Raison sociale
  if LeType = '001' then
  begin
    if LeMode = '003' then //mixte
      RTLanceFiche_RTTOPCIBLE('RT', 'RTTOPCIBLE', '', '', LeCode+';'+LeMode);
    if LeMode = '001' then //Tiers
      RTLanceFiche_RTTOPTIERS('RT', 'RTTOPTIERS', '', '', LeCode+';'+LeMode);
    if LeMode = '002' then //Suspect
      RTLanceFiche_RTCIBLESUS('RT', 'RTCIBLESUS', '', '', LeCode+';'+LeMode);
  end
  else
  begin
    if LeMode = '003' then  // mixte
      RTLanceFiche_UTOFRTTOPCONTACT('RT', 'RTTOPCONTACT', '', '', LeCode+';'+LeMode);
    if LeMode = '001' then // tiers
      RTLanceFiche_RTTOPTIERSCON('RT', 'RTTOPTIERSCON', '', '', LeCode+';'+LeMode);
    if LeMode = '002' then // suspect
      RTLanceFiche_RTTOPSUSCON('RT', 'RTTOPSUSCON', '', '', LeCode+';'+LeMode);
  end;

  TFMul(ecran).BChercheClick(Ecran);
end;




procedure TOF_UTOFCIBLAGE.LaListeOnFlipSelection(Sender: TObject);
var
   Qr : TQuery;
   SInfo : string;
   NbEnreg : integer;
   F : TFmul;
begin

   if GetField('RCB_CODECIBLAGE')=null then
      exit;

   SInfo := GetField('RCB_CODECIBLAGE')+' : ';
   Qr := OpenSql ('SELECT COUNT(*) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+
      '" AND RVB_VISA="002"', True);
   if not Qr.Eof then
      NbEnreg := Qr.Fields[0].AsInteger
   else
      NbEnreg := 0;
   SInfo := SInfo+IntToStr(NbEnreg)+' Visa / ';
   Ferme(Qr);

   Qr := OpenSql ('SELECT COUNT(*) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+
      '" AND RVB_VISA="001"', True);
   if not Qr.Eof then
      NbEnreg := Qr.Fields[0].AsInteger
   else
      NbEnreg := 0;
   SInfo := SInfo+IntToStr(NbEnreg)+' PréVisa / ';
   Ferme(Qr);

   Qr := OpenSql ('SELECT COUNT(*) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+
      '" AND RVB_VISA="003"', True);
   if not Qr.Eof then
      NbEnreg := Qr.Fields[0].AsInteger
   else
      NbEnreg := 0;
   SInfo := SInfo+IntToStr(NbEnreg)+' Non pris';
   Ferme(Qr);

   F := TFmul (ecran);
   F.FListe.Hint := SInfo;

end;

procedure TOF_UTOFCIBLAGE.LaListeOnDblClick(Sender: TObject);
begin
   if GetField('RCB_CODECIBLAGE')=null then exit;

   AglLanceFiche('RT', 'RTSAISIECIBLAGE', '', GetField('RCB_CODECIBLAGE'), 'MONOFICHE;ACTION=MODIFICATION;BLOQUE=' + BoolToStr_(Bbloque) + ';AVERTISSEMENT=' + BoolToStr_(BAvert));
//   TFmul(ecran).Q.UpdateCriteres;
TFMul(ecran).BChercheClick(Ecran);

end;

procedure TOF_UTOFCIBLAGE.ZoomOnClick(Sender: TObject);
begin
   if GetField('RCB_CODECIBLAGE')=null then exit;

   AglLanceFiche('RT', 'RTSAISIECIBLAGE', '', GetField('RCB_CODECIBLAGE'), 'MONOFICHE;ACTION=CONSULTATION;BLOQUE=' + BoolToStr_(Bbloque) + ';AVERTISSEMENT=' + BoolToStr_(BAvert));

end;

procedure TOF_UTOFCIBLAGE.BinsertOnClick(Sender: TObject);
//var
//  StrMsg                : String;
//  bBloque               : boolean;

begin
  //FQ 10484
(*
  StrMsg                := JaiLeDroitCiblage(bBloque);
  if StrMsg <> '' then
  begin
    StrMsg              := 'Attention !' + StrMsg;
    PGIBox(StrMsg, Ecran.Caption);
  end;
*)

  AglLanceFiche('RT', 'RTSAISIECIBLAGE', '', '', 'MONOFICHE;ACTION=CREATION;BLOQUE=' + BoolToStr_(Bbloque) + ';AVERTISSEMENT=' + BoolToStr_(BAvert));
  TFMul(ecran).BChercheClick(Ecran);


end;



procedure TOF_UTOFCIBLAGE.BsupprimerOnClick(Sender: Tobject);
var
  F : Tfmul;
  StrWhere  : String;
  StrSql1  : String;
  StrSql2  : String;

begin
  F := TFMul(Ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
       PGIInfo('Aucun élément sélectionné');
       exit;
  end;

  if PGIAsk('Attention !#10#13'+
            'Cette commande va supprimer tous les ciblages sélectionnés.#10#13'+
            'Confirmez-vous le traitement ?', Ecran.caption) = mrNo then
            Exit;

  InitMoveProgressForm(nil,'Suppression en cours','',3,True,True);
  StrWhere := MakeWhere;
  BEGINTRANS;
  Try
    if StrWhere <> '' then
    begin
      StrSql1  := 'DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE IN (SELECT RCB_CODECIBLAGE FROM CIBLAGE WHERE '+StrWhere+')';
      StrSql2  := 'DELETE FROM CIBLAGE WHERE ' +StrWhere;
    end
    else
    begin
      StrSql1  := 'DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE IN (SELECT RCB_CODECIBLAGE FROM CIBLAGE)';
      StrSql2  := 'DELETE FROM CIBLAGE';
    end;

    MoveCurProgressForm('Suppression des éléments du/des ciblage(s)');
    ExecuteSQL(StrSql1);
    MoveCurProgressForm('Suppression du/des ciblage(s)');
    ExecuteSQL(StrSql2);
    COMMITTRANS;
  except
    on E:Exception do
    begin
      ROLLBACK;
      PGIInfo('Erreur : ' + E.message); //mcd 28/10/2008 mis après rollback pour eviter pb lock
    end;
  end;

  FiniMoveProgressForm;
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;


end;


procedure TOF_UTOFCIBLAGE.RbOnClick(Sender: Tobject);
var
  StrRb                 : String;

begin
  With Sender as TRadioButton do
  begin
    StrRb               := Name;
    if StrRb = 'RB1' then
      SetControlText('RCB_MODECIBLAGE', '001')
    else if StrRb = 'RB2' then
      SetControlText('RCB_MODECIBLAGE', '002')
    else
      SetControlText('RCB_MODECIBLAGE', '003');
  end;

  TFMul(ecran).BChercheClick(Ecran);
end;



Initialization
  registerclasses ( [ TOF_UTOFCIBLAGE ] ) ;
end.
