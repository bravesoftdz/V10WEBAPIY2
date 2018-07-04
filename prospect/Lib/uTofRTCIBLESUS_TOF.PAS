{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/06/2006
Modifié le ... : 08/06/2007
Description .. : Source TOF de la FICHE : RTCIBLESUS ()
Suite ........ : CIBLAGE SUSPECT - RAISON SOCIALE
Mots clefs ... : TOF;RTCIBLESUS
*****************************************************************}
Unit uTofRTCIBLESUS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     uTob,
     MaineAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UtilGC,
     HTB97,
     KPMGUtil,
     ed_tools,
     UtilRT,
     UTOF ;

Type
  TOF_RTCIBLESUS = Class (TOF)
  private
    CodeCiblage : String;
    ModeCiblage : String;

  public

    BAjoute : TToolBarButton97;
    BEnleve : TToolBarButton97;
    BZoom : TToolbarButton97;
    LaListe : THGrid;
    SqlCible            : String;
    SqlNonCible         : String;

    procedure BAjouteOnClick ( Sender : TObject );
    procedure BEnleveOnClick ( Sender : TObject );
    procedure LaListeOnDblClick ( Sender : TObject );
    procedure BZoomOnclick ( Sender : Tobject );

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_RTCIBLESUS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RTCIBLESUS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTCIBLESUS.OnArgument (Argument : String ) ;
var
   critere : string;
   F : TFMul;
   StrWhere             : String;

begin
  Inherited ;
  F := TFMul(ecran);

  SetControlVisible('BOUVRIR', False);

  BAjoute := TToolBarButton97(getcontrol('BAJOUTE'));
  BAjoute.OnClick := BAjouteOnClick;

  BEnleve := TToolBarButton97(GetControl('BENLEVE'));
  BEnleve.OnClick := BEnleveOnClick;

  LaListe := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick := LaListeOnDblClick;

  BZoom := TToolbarButton97(GetControl('BZOOM'));
  if Assigned (BZoom) then
  begin
    BZoom.Visible := True;
    BZoom.OnClick := BZoomOnclick;
  end;


  //Mise à jour des champs libres
  RTMajChampsLibresSuspect(F);


  Critere := UpperCase(Trim(Argument));
  CodeCiblage           := READTOKENST(Critere);
  ModeCiblage           := READTOKENST(Critere);

  F.caption := 'Sélection des ciblages : '+codeciblage;
  UpdateCaption(ecran);

  SetControlVisible('BINSERT', False);
  SetControlVisible('BSELECTALL', True);

  SqlCible              := 'VCM_CODECIBLAGE="' + CodeCiblage + '" ';
  SqlNonCible           := '(RSU_SUSPECT NOT IN (SELECT RVB_SUSPECT FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '" ) ' +
                           'AND VCM_CODECIBLAGE="")';

  StrWhere              := SqlCible + 'OR ' + SqlNonCible;

  SetControlText ('XX_WHERE', StrWhere);

  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('RSU_REPRESENTANT', False);
    SetControlVisible ('TRSU_REPRESENTANT', False);
  end;

end ;

procedure TOF_RTCIBLESUS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTCIBLESUS.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_RTCIBLESUS.OnUpdate ;
begin
   Inherited
end ;

procedure TOF_RTCIBLESUS.OnLoad ;
var
//   Lreq : String;
   F : TFMul;
   StrWhere             : String;


begin
  Inherited ;

  F := TFMul(ecran);

//  Lreq := ConstruitRequete ('VCM_COCHE', 'VCM_CODECIBLAGE', F);
  ReConstruireListe(F);

  if GetCheckBoxState('COCHEE') = CbGrayed then
     StrWhere           := SqlCible + ' OR ' + SqlNonCible
  else if GetCheckBoxState('COCHEE') = CbChecked then
     StrWhere           := SqlCible
  else
     StrWhere           := SqlNonCible;

   SetControlText ('XX_WHERE', StrWhere);


  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('LISTE');
  {$ENDIF}

  //recherche le nombre d'éléments ciblés
  if GetControl('NBELEMENTS')<>nil then
    SetControlText('NBELEMENTS', IntToStr(TrouveNbElement(CodeCiblage)));

end ;

procedure TOF_RTCIBLESUS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTCIBLESUS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTCIBLESUS.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/06/2006
Modifié le ... :   /  /
Description .. : ajout d'éléments dans le ciblage
Mots clefs ... :
*****************************************************************}
procedure TOF_RTCIBLESUS.BAjouteOnClick(Sender: TObject);
var
   F : TFmul;
   LgFiche : integer;

begin

  F := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
       messagealerte('Aucun élément sélectionné');
       exit;
  end;


  //ecriture des previsa
  if F.Fliste.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not F.FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
  {$ENDIF}
     F.bSelectAllClick(nil);
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.Q.recordcount,True,True);
     F.Q.First;
     While not F.Q.Eof do
     begin
        CreationVisa (CodeCiblage, F.Q.FindField('RSU_SUSPECT').AsString, 'SUS', '001', '0');
        MoveCurProgressForm('');
        F.Q.Next;
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.FListe.nbSelected-1,True,True);
     for lgfiche := 0 to F.Fliste.NbSelected - 1 do
     begin
        F.Fliste.GotoLeBookmark(lgfiche);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}
        CreationVisa (CodeCiblage, F.Q.FindField('RSU_SUSPECT').AsString, 'SUS', '001', '0');
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
Créé le ...... : 23/06/2006
Modifié le ... :   /  /
Description .. : suppression d'éléments dans le ciblage
Mots clefs ... :
*****************************************************************}
procedure TOF_RTCIBLESUS.BEnleveOnClick(Sender: TObject);
var
   F : TFmul;
   LgFiche : integer;
begin
  F := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
       messagealerte('Aucun élément sélectionné');
       exit;
  end;


  //Suppession des ciblages
  if F.Fliste.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not F.FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
  {$ENDIF}
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.Q.recordcount,True,True);
     F.Q.First;
     While not F.Q.Eof do
     begin
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('RSU_SUSPECT').AsString+'"');

        F.Q.Next;
        MoveCurProgressForm('');
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.FListe.nbSelected-1,True,True);
     for lgfiche := 0 to F.Fliste.NbSelected - 1 do
     begin
        F.Fliste.GotoLeBookmark(lgfiche);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('RSU_SUSPECT').AsString+'"');
        MoveCurProgressForm('');
     end;
  end;


  FiniMoveProgressForm;
  //F.Q.UpdateCriteres;
  F.BChercheClick(Ecran);

  F.FListe.ClearSelected;
end;


procedure TOF_RTCIBLESUS.LaListeOnDblClick(Sender: TObject);
begin
  if VarIsNull(GetField('RSU_SUSPECT')) then
     exit;

  if GetField('VCM_CODECIBLAGE') <> '' then
     ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                '" AND RVB_SUSPECT="'+GetField('RSU_SUSPECT')+'"')
  else
     CreationVisa(CodeCiblage, Getfield('RSU_SUSPECT'), 'SUS', '001', '0');

  TFMul(ecran).BChercheClick(Ecran);

end;


procedure TOF_RTCIBLESUS.BZoomOnclick(Sender: Tobject);
begin
  if Getfield('RSU_SUSPECT')=Null then
    exit;
  AGLLanceFiche('RT', 'RTSUSPECTS', '', GetField('RSU_SUSPECT'), 'MONOFICHE;ACTION=CONSULTATION');
end;


Initialization
  registerclasses ( [ TOF_RTCIBLESUS ] ) ;
end.
