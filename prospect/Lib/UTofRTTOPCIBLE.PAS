{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 16/05/2006
Modifié le ... : 08/06/2007
Description .. : Source TOF de la FICHE : RTTOPCIBLE ()
Suite ........ : Sélection et ajout de prévisa
Suite ........ : CIBLAGE MIXTE - RAISON SOCIALE
Mots clefs ... : TOF;RTTOPCIBLE
*****************************************************************}
Unit UTofRTTOPCIBLE ;

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
     eMul,
     Maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,
     UTOF,
     HTB97,
     Graphics,
     ed_tools,
     utilGC,
     KPMGUtil,
     uTofMixte,
     UtilPGI;

Type
  TOF_RTTOPCIBLE = Class (TOF_MIXTE)
   private
    CodeCiblage         : String;
    ModeCiblage         : String;

   public
    BAjoute             : TToolBarButton97;
    BEnleve             : TToolBarButton97;
    LaListe             : THGrid;
    SqlCible            : String;
    SqlNonCible         : String;
    ConfWhere           : String;

    procedure BAjouteOnClick ( Sender : TObject );
    procedure BEnleveOnClick ( Sender : TObject );
    procedure LaListeOnDblClick ( Sender : TObject );

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_RTTOPCIBLE(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RTTOPCIBLE(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_RTTOPCIBLE.OnArgument (Argument : String ) ;
var
  F                     : TFmul;
  Critere               : string;

begin
  With ParamMixte do
  begin
    Suffixe             := 'VCM';
    Action              := taConsult;
    TypeFiche           := 'CIBLAGE';
  end;

  Inherited ;

  F                     := TFmul (ecran);
  SetControlVisible('BOUVRIR', False);

  BAjoute               := TToolBarButton97(getcontrol('BAJOUTE'));
  BAjoute.OnClick       := BAjouteOnClick;

  BEnleve               := TToolBarButton97(GetControl('BENLEVE'));
  BEnleve.OnClick       := BEnleveOnClick;

  LaListe               := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick    := LaListeOnDblClick;

  SetControlVisible('BAJOUTE', True);
  SetControlVisible('BENLEVE', True);

  Critere               := UpperCase(Trim(Argument));
  CodeCiblage           := READTOKENST(Critere);
  ModeCiblage           := READTOKENST(Critere);

  SqlCible              := 'VCM_CODECIBLAGE="' + CodeCiblage + '" ';
  SqlNonCible           := '((VCM_CODE NOT IN (SELECT RVB_TIERS FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '") ' +
                           'AND VCM_CODECIBLAGE="" AND VCM_NATURE<>"SUS") ' +
                           'OR (VCM_CODE NOT IN (SELECT RVB_SUSPECT FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '") ' +
                           'AND VCM_CODECIBLAGE="" AND VCM_NATURE="SUS"))';

  ConfWhere             := GetControlText('XX_WHERE');    // on récupère la restriction des fiches tiers
  SetControlVisible ('BSelectAll', True);

  F.Caption             := 'Sélection des ciblages : '+CodeCiblage;
  UpdateCaption(ecran);
{$Ifdef GIGI}
 if (GetControl('VCM_REPRESENTANT') <> nil) then  SetControlVisible('VCM_REPRESENTANT',false);
 if (GetControl('TVCM_REPRESENTANT') <> nil) then  SetControlVisible('TVCM_REPRESENTANT',false);
 if (GetControl('VCM_ZONECOM') <> nil) then  SetControlVisible('VCM_ZONECOM',false);
 if (GetControl('TVCM_ZONECOM') <> nil) then  SetControlVisible('TVCM_ZONECOM',false);
{$endif GIGI}


end ;

procedure TOF_RTTOPCIBLE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPCIBLE.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_RTTOPCIBLE.OnUpdate ;
begin
  Inherited ;

end;



procedure TOF_RTTOPCIBLE.OnLoad ;
var
   StrWhere             : String;
   F                    : TFMul;

begin
  Inherited ;

  F                     := TFMul(ecran);

  ReConstruireListe(F);

  if GetCheckBoxState('COCHEE') = CbGrayed then
     StrWhere           := SqlCible + ' OR ' + SqlNonCible
  else if GetCheckBoxState('COCHEE') = CbChecked then
     StrWhere           := SqlCible
  else
     StrWhere           := SqlNonCible;

  SetControlText('XX_WHERE', StrWhere);
  SetControlText('XX_WHERESPE', ConfWhere);
//   UpdateWhere(StrWhere);


  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('LISTE');
  {$ENDIF}

  if GetControl('NBELEMENTS')<>nil then
    SetControlText('NBELEMENTS', IntToStr(TrouveNbElement(CodeCiblage)));

end ;



procedure TOF_RTTOPCIBLE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPCIBLE.OnDisplay () ;
begin
  Inherited ;

end ;


procedure TOF_RTTOPCIBLE.OnCancel () ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 01/06/2006
Modifié le ... :   /  /
Description .. : bouton permettant d'ajouter des ciblages
Mots clefs ... :
*****************************************************************}
procedure TOF_RTTOPCIBLE.BAjouteOnClick(Sender: TObject);
var
   F : TFmul;
   LgFiche : integer;
 //fonction enregistrement des visa

begin
  F                     := Tfmul (ecran);
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
        CreationVisa (CodeCiblage, F.Q.FindField('VCM_CODE').AsString, F.Q.FindField('VCM_NATURE').AsString, '001', '0');
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
        CreationVisa (CodeCiblage, F.Q.FindField('VCM_CODE').AsString, F.Q.FindField('VCM_NATURE').AsString, '001', '0');
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
Créé le ...... : 01/06/2006
Modifié le ... :   /  /
Description .. : Bouton permettant d'enlever des ciblages
Mots clefs ... :
*****************************************************************}
procedure TOF_RTTOPCIBLE.BEnleveOnClick(Sender: TObject);
var
   F                    : TFmul;
   LgFiche              : integer;

begin
  F                     := Tfmul (ecran);
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
     InitMoveProgressForm(nil,'Traitement en cours','',F.Q.recordcount,True,True);
     F.Q.First;
     While not F.Q.Eof do
     begin
        if F.Q.FindField('VCM_NATURE').AsString='SUS' then
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('VCM_CODE').AsString+'"')
        else
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_TIERS="'+F.Q.FindField('VCM_CODE').AsString+'"');

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
        if F.Q.FindField('VCM_NATURE').AsString='SUS' then
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('VCM_CODE').AsString+'"')
        else
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_TIERS="'+F.Q.FindField('VCM_CODE').AsString+'"');
        MoveCurProgressForm('');

     end;
  end;

  FiniMoveProgressForm;
  //F.Q.UpdateCriteres;
  F.BChercheClick(Ecran);

  F.FListe.ClearSelected;


end;



procedure TOF_RTTOPCIBLE.LaListeOnDblClick(Sender: TObject);
begin
  if VarIsNull(GetField('VCM_CODE')) then
     exit;

  if GetField('VCM_CODECIBLAGE') <> '' then
  begin
     if GetField('VCM_NATURE') = 'SUS' then
        ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                   '" AND RVB_SUSPECT="'+GetField('VCM_CODE')+'"')
     else
        ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                   '" AND RVB_TIERS="'+GetField('VCM_CODE')+'"');
  end
  else
     CreationVisa (CodeCiblage, GetField('VCM_CODE'), GetField('VCM_NATURE'), '001', '0');

  TFMul(ecran).BChercheClick(Ecran);

end;



Initialization
  registerclasses ( [ TOF_RTTOPCIBLE ] ) ;
end.
