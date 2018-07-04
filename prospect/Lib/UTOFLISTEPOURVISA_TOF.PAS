{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 18/05/2006
Modifié le ... : 22/05/2006
Description .. : Source TOF de la FICHE : UTOFLISTEPOURVISA ()
Suite ........ : Sélection des cibles pour mettre le visa
Suite ........ : Si suspect alors transfert vers prospect
Mots clefs ... : TOF;UTOFLISTEPOURVISA
*****************************************************************}
Unit UTOFLISTEPOURVISA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     Maineagl,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     ed_tools,
     paramsoc,
     ent1,
     KPMGUtil ;

Type
  TOF_UTOFLISTEPOURVISA = Class (TOF)
   public
    CodeCiblage : String;

    LaListe : THGrid;
    BVisa : TToolBarButton97;
    BPreVisa : TToolBarButton97;
    BNonPris : TToolBarButton97;
    Bselect : TToolBarButton97;
    BCible : TToolBarButton97;
    ModeCiblage : String;

    procedure LaListeOnDblClick ( Sender : TObject );
    procedure BVisaOnclick ( Sender : TObject );
    procedure BPreVisaOnclick ( Sender : TObject );
    procedure BNonPrisOnclick ( Sender : TObject );
    procedure GestionVisa (visa : String);
    procedure BselectOnClick ( Sender : TObject );
    procedure BCibleOnClick ( Sender : TObject );



    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_utoflistepourvisa(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation
uses UFonctionsCBP;

Function RTLanceFiche_utoflistepourvisa(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_UTOFLISTEPOURVISA.OnArgument (Arguments : String ) ;
var
   ReqListe : String;
   Q : TQuery;
   Cbour : String;
   StrMsg : String;
   BloqueTrans : boolean;
   LectureSeule : Boolean;

begin
  Inherited ;
  CodeCiblage := Arguments;

//si ciblage déjà validé, on autorise la consultation uniquement
  Q := OpenSQL('SELECT RCB_VALIDATION, RCB_MODECIBLAGE FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+CodeCiblage+'" ', True);
  //si déjà validé
  LectureSeule := False;
  if Q.Fields[0].Asstring='X' then
  begin
    LectureSeule := True;
    SetControlVisible('BPREVISA', False);
    SetControlVisible('BVISA', False);
    SetControlVisible('BNONPRIS', False);
    SetControlVisible('BSELECT', False);
    SetControlVisible('BCIBLE', False);
    PGIInfo  ('Ce ciblage est déjà validé, vous ne pouvez que le consulter.');
  end;
  Modeciblage := Q.Fields[1].asstring;
  Ferme(Q);

  ReqListe := 'VVM_CODECIBLAGE="'+CodeCiblage+'"';
  SetControlText ('XX_WHERE', ReqListe);

  SetControlVisible ('BSelectAll', True);

  TFMul(ecran).Caption := 'Visa sur le ciblage '+CodeCiblage;


  LaListe := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick := LaListeOnDblClick;

  SetControlVisible('BOUVRIR', False);

  BVisa := TToolBarButton97(GetControl('BVISA'));
  BVisa.OnClick := BVisaOnClick;
  BPreVisa := TToolBarButton97(GetControl('BPREVISA'));
  BPreVisa.OnClick := BPreVisaOnClick;
  BNonPris := TToolBarButton97(GetControl('BNONPRIS'));
  BNonPris.OnClick := BNonPrisOnClick;
  Bselect := TToolBarButton97(GetControl('Bselect'));
  Bselect.OnClick := BselectOnClick;
  BCible := TToolBarButton97(GetControl('BCIBLE'));
  BCible.OnClick := BCibleOnClick;

  if ModeCiblage='001' then   //tiers
    SetControlVisible('Bselect', True);

  StrMsg := 'Attention';
  BloqueTrans := False;

  if (Modeciblage<>'001') and (GetParamSocSecur('SO_GCNUMTIERSAUTO', False, True) = False) then
  begin
//    PGIInfo('Vous N''êtes pas en numérotation automatique des tiers.#10#13'+
//            'La commande de validation du ciblage est bloquée.', 'Avertissement');
    StrMsg := StrMsg + '#10#13  - Vous n''êtes pas en numérotation automatique des tiers.';
    SetControlVisible('Bselect', False);
    BloqueTrans := True;
  end;

  Cbour := GetParamSocSecur('SO_BOURREAUX','X',True);
  Cbour := UpperCase_(Cbour);
  if (not (Cbour[1] in ['A'..'Z'])) and (ModeCiblage<>'001') then
  begin
//    PGIInfo('Attention !#10#13Le caractère de bourrage du compte auxiliaire est numérique.#10#13'+
//            'Le transfert de suspect en prospect est impossible.', 'Avertissement');
    StrMsg := StrMsg + '#10#13  - Le caractère de bourrage du compte auxiliaire est numérique.';
    SetControlVisible('Bselect', False);
    BloqueTrans := True;
  end;

  //FQ 10479
  if GetParamSocSecur('SO_GCENSEIGNETAB', False, True) then
  begin
//   PGIInfo('Vous avez activé la Gestion des enseignes en tablette dans le module Paramètres#10#13'+
//            'il n''est donc pas possible de transférer les enseignes présentes #10#13'+
//            'sur les fiches suspects si celles-ci figurent en tant que texte libre.', 'Avertissement');
    StrMsg := StrMsg + '#10#13  - Vous avez activé la gestion des enseignes en tablette dans le module Paramètres#10#13'+
                       '    il n''est donc pas possible de transférer les enseignes présentes #10#13'+
                       '    sur les fiches suspects si celles-ci figurent en tant que texte libre.';
    BloqueTrans := True;
  end;
  if (BloqueTrans) and (not LectureSeule) then
    PGIInfo(StrMsg, 'Avertissement');

  // PL le 16/04/07 : j'ajoute la condition sur le Mode de ciblage car
  // on peut faire un ciblage client/prospect si on n'a pas le concept de transfert par lot des
  // suspects en prospect    
  if (not ExJaiLeDroitConcept(TConcept(gcTransfertSusPro), False))  and (ModeCiblage<>'001') then
  begin
    SetControlVisible('BSELECT', False);
  end;

end ;


procedure TOF_UTOFLISTEPOURVISA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_UTOFLISTEPOURVISA.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFLISTEPOURVISA.LaListeOnDblClick(Sender: TObject);
begin
if GetField('VVM_CODE')='' then
   exit;

//on appelle la fiche du tiers ou du suspect pour informations


if GetField('VVM_NATURE')='SUS' then
   AGLLanceFiche ('RT', 'RTSUSPECTS', '', GetField('VVM_CODE'), 'MONOFICHE;ACTION=CONSULTATION')
else if GetField('VVM_NATURE')='CLI' then
{CRM_20080722_CD_012;10802_DEB}
//   AGLLanceFiche('GC', 'GCTIERS','',GetField('VVM_CODE'),'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=CLI')
   V_PGI.DispatchTT(8,taConsult, GetField('VVM_CODE'), '', 'MONOFICHE;T_NATUREAUXI=CLI')
{CRM_20080722_CD_012;10802_FIN}
else if GetField('VVM_NATURE')='PRO' then
{CRM_20080722_CD_012;10802_DEB}
//   AGLLanceFiche('GC', 'GCTIERS','',GetField('VVM_CODE'),'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=PRO');
   V_PGI.DispatchTT(8,taConsult, GetField('VVM_CODE'), '', 'MONOFICHE;T_NATUREAUXI=PRO')
{CRM_20080722_CD_012;10802_FIN}



//TFmul(ecran).Q.UpdateCriteres;

end;



procedure TOF_UTOFLISTEPOURVISA.BNonPrisOnclick(Sender: TObject);
begin
   GestionVisa('003');
end;

procedure TOF_UTOFLISTEPOURVISA.BPreVisaOnclick(Sender: TObject);
begin
   GestionVisa('001');
end;

procedure TOF_UTOFLISTEPOURVISA.BVisaOnclick(Sender: TObject);
begin
   GestionVisa('002');
end;

procedure TOF_UTOFLISTEPOURVISA.GestionVisa(visa: String);
var
   F : TFmul;
   LgFiche : integer;

   //Update de la table AFVCIBLAGE avec le type de visa sélectionné
   Function MajVisa(Clecible, naturecible : String) : Boolean;
   begin

   if naturecible='SUS' then
      ExecuteSql ('UPDATE CIBLAGEELEMENT SET RVB_VISA="'+Visa+'" WHERE RVB_CODECIBLAGE="'+CodeCiblage+
      '" AND RVB_TIERS="XXX" AND RVB_CONTACT="XXX" AND RVB_SUSPECT="'+Clecible+'"')
   else
      ExecuteSql ('UPDATE CIBLAGEELEMENT SET RVB_VISA="'+Visa+'" WHERE RVB_CODECIBLAGE="'+CodeCiblage+
      '" AND RVB_TIERS="'+Clecible+'" AND RVB_CONTACT="XXX" AND RVB_SUSPECT="XXX"');


   Result := True;

   end;


begin
F := TFmul(ecran);

//Traitement sur la sélection
if (F.FListe.nbSelected=0) and (not F.FListe.AllSelected) then
begin
     messagealerte('Aucun élément sélectionné');
     exit;
end;

if F.FListe.AllSelected then
begin
{$IFDEF EAGLCLIENT}
    if not F.FetchLesTous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
   F.Q.First;
   InitMoveProgressForm(nil,'Traitement en cours ...','',F.Q.recordcount,True,True);
   while not F.Q.Eof do
   begin
      MajVisa(F.Q.FindField('VVM_CODE').AsString, F.Q.FindField('VVM_NATURE').AsString);

      MoveCurProgressForm('');
      F.Q.Next;
   end;
end
else
begin
   InitMoveProgressForm(nil,'Traitement en cours','',F.FListe.nbSelected-1,True,True);
   for LgFiche := 0 to F.FListe.nbSelected-1 do
   begin
      F.Fliste.GotoLeBookmark(lgfiche);
      {$IFDEF EAGLCLIENT}
      F.Q.TQ.Seek(F.FListe.row -1);
      {$ENDIF}
      MajVisa(F.Q.FindField('VVM_CODE').AsString, F.Q.FindField('VVM_NATURE').AsString);
      MoveCurProgressForm('');

   end;
end;
FiniMoveProgressForm;
//F.Q.UpdateCriteres;
F.BChercheClick(Ecran);
F.FListe.ClearSelected;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/05/2006
Modifié le ... :   /  /
Description .. : bouton validation du ciblage
Mots clefs ... :
*****************************************************************}
procedure TOF_UTOFLISTEPOURVISA.BselectOnClick(Sender: TObject);
var
   Qr : TQuery;
   Qr2 : TQuery;
   NbEnreg : integer;
   CodeTiers : String;
   GereEnseigne : Boolean;
   BlValide : boolean;

begin
NbEnreg := 0;
// Recherche s'il reste des Prévisas
Qr := OpenSQL ('SELECT COUNT(*) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'" AND RVB_VISA="001"', True);
If not Qr.Eof then
   NbEnreg := Qr.Fields[0].AsInteger;
Ferme(Qr);

If NbEnreg<>0 then
begin
   MessageAlerte('Il reste des Prévisa');
   exit;
end;

// Recherche s'il y a des visas
NbEnreg := 0;
Qr := OpenSQL ('SELECT COUNT(*) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'" AND RVB_VISA="002"', True);

If not Qr.Eof then
   NbEnreg := Qr.Fields[0].AsInteger;
Ferme(Qr);

If NbEnreg=0 then
begin
   MessageAlerte ('Il n''y a rien à valider');
   exit;
end;

TFmul(Ecran).Cursor := crHourGlass;

//FQ 10479
GereEnseigne := GetParamSocSecur('SO_GCENSEIGNETAB', False, True);
//Recheche des suspects pour le passage en prospect
Qr := OpenSQL ('SELECT CIBLAGEELEMENT.*,RSU_LIBELLE FROM CIBLAGEELEMENT,SUSPECTS WHERE RVB_CODECIBLAGE="'+CodeCiblage+
      '" AND RVB_TIERS="XXX" AND RVB_CONTACT="XXX" AND RVB_SUSPECT<>"" AND RVB_NUMEROCONTACT="0" AND RSU_SUSPECT=RVB_SUSPECT' +
      ' AND RVB_VISA<>"003"', True);
Qr.First;
InitMoveProgressForm(nil,'Traitement en cours ...','',Qr.RecordCount,True,True);

BlValide := true;
While not Qr.Eof do
begin
  BEGINTRANS;
  Try
     //passage suspect vers prospect
     MoveCurProgressForm('Transfert suspect : '+Qr.findfield('RSU_LIBELLE').AsString);
     CodeTiers := SuspectVersProspect(Qr.findfield('RVB_SUSPECT').AsString, True, GereEnseigne);
     //mise à jour de AFVCIBLAGE de l'ex suspect afin de le retrouver dans toute la table
     If CodeTiers<>'' then
     begin
        Qr2 := OpenSql('SELECT * FROM TIERS WHERE T_TIERS="'+CodeTiers+'"', True);

        ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_CONTACT="'+Qr2.FindField('T_AUXILIAIRE').AsString+
           '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+Qr.findfield('RVB_SUSPECT').AsString+
           '" AND RVB_NUMEROCONTACT="1"');

        ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_TIERS="'+CodeTiers+
           '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+Qr.findfield('RVB_SUSPECT').AsString+
           '" AND RVB_NUMEROCONTACT="0"');


        ferme(Qr2);
     end;
    COMMITTRANS;
  except
    ROLLBACK;
    BlValide := false;
  end;

   Qr.Next;

end;


FiniMoveProgressForm;

//verrou du ciblage
if BlValide then
begin
  ExecuteSQL ('UPDATE CIBLAGE SET RCB_VALIDATION="X" WHERE RCB_CODECIBLAGE="'+CodeCiblage+'"');
  PGIInfo ('Ciblage Validé');
end
else
  PGIInfo('Problèmes rencontrés durant le transfert des suspects');

//TFmul(ecran).Q.UpdateCriteres;
TFMul(ecran).BChercheClick(Ecran);
Ferme(Qr);
TFmul(Ecran).Cursor := crDefault;




TFMul(ecran).Close;

end;



procedure TOF_UTOFLISTEPOURVISA.BCibleOnClick(Sender: TObject);
begin
   if ModeCiblage = '003' then
      AglLanceFiche('RT', 'RTTOPCIBLE', '', '', CodeCiblage+';'+ModeCiblage);
   if ModeCiblage = '001' then
      AglLanceFiche('RT', 'RTTOPTIERS', '', '', CodeCiblage+';'+ModeCiblage);
   if ModeCiblage = '002' then
      AglLanceFiche('RT', 'RTCIBLESUS', '', '', CodeCiblage+';'+ModeCiblage);
   TFMul(ecran).BChercheClick(ecran);
end;




Initialization
  registerclasses ( [ TOF_UTOFLISTEPOURVISA ] ) ;
end.
