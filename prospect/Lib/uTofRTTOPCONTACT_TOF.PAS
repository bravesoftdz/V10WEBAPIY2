{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 23/05/2006
Modifié le ... : 08/06/2007
Description .. : Source TOF de la FICHE : UTOFRTTOPCONTACT ()
Suite ........ : Topage des contacts suspects et tiers
Suite ........ : CIBLAGE MIXTE - CONTACT
Mots clefs ... : TOF;UTOFRTTOPCONTACT
*****************************************************************}
Unit uTofRTTOPCONTACT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,Fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,Maineagl,
 //    uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     utob,
     UtilGc,
     ed_tools,
     KPMGUtil,
     uTofMixte,
     HTB97
 ;

Type
  TOF_UTOFRTTOPCONTACT = Class (TOF_MIXTE)
   private
      CodeCiblage       : String;
      ModeCiblage       : String;
      BAJOUTE           : TToolbarButton97;
      BENLEVE           : TToolBarButton97;
      LaListe           : THGrid;
      SqlCible          : String;
      SqlNonCible       : String;
      ConfWhere         : String;

      procedure BAjouteOnClick ( Sender : TObject );
      procedure BenleveOnclick ( Sender : TObject );
      procedure LaListeOnDblClick ( Sender : TObject );


   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_UTOFRTTOPCONTACT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_UTOFRTTOPCONTACT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_UTOFRTTOPCONTACT.OnArgument (Argument : String ) ;
var
  Critere               : string;
  F                     : TFMul;

begin
  With ParamMixte do
  begin
    Suffixe             := 'CCM';
    Action              := taConsult;
    TypeFiche           := 'CIBLAGE';
  end;


  Inherited ;

  SetControlVisible('BOUVRIR', False);

  BAjoute               := TToolBarButton97(GetControl('BAJOUTE'));
  BAjoute.OnClick       := BAjouteOnClick;

  BEnleve               := TToolBarButton97(GetControl('BENLEVE'));
  BEnleve.OnClick       := BEnleveOnCLick;

  LaListe               := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick    := LaListeOnDblClick;

  Critere               := UpperCase(Trim(Argument));
  CodeCiblage           := READTOKENST(Critere);
  ModeCiblage           := READTOKENST(Critere);

  F                     := TFMul(ecran);

  SqlCible              := 'CCM_CODECIBLAGE="' + CodeCiblage + '" ';
  SqlNonCible           := '((CCM_NUMCONTACT NOT IN (SELECT RVB_CONTACT || ";" || CAST(RVB_NUMEROCONTACT AS VARCHAR(5)) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '") ' +
                           'AND CCM_CODECIBLAGE="" AND CCM_NATURE<>"SUS") ' +
                           'OR (CCM_NUMCONTACT NOT IN (SELECT RVB_SUSPECT || ";1" FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '") ' +
                           'AND CCM_CODECIBLAGE="" AND CCM_NATURE="SUS"))';

  ConfWhere             := GetControlText('XX_WHERE');      // on récupère les restrictions des fiches tiers
  SetControlVisible ('BSelectAll', True);

  F.Caption := 'Sélection des ciblages : '+ CodeCiblage;
  UpdateCaption(Ecran);

end ;



procedure TOF_UTOFRTTOPCONTACT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFRTTOPCONTACT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFRTTOPCONTACT.OnUpdate ;
begin
  Inherited ;

;

end ;

procedure TOF_UTOFRTTOPCONTACT.OnLoad ;
var
  F                     : TFMul;
  StrWhere              : String;

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

  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('LISTE');
  {$ENDIF}

  if GetControl('NBELEMENTS')<>nil then
    SetControlText('NBELEMENTS', IntToStr(TrouveNbElement(CodeCiblage)));

end ;


procedure TOF_UTOFRTTOPCONTACT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFRTTOPCONTACT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFRTTOPCONTACT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFRTTOPCONTACT.BAjouteOnClick(Sender: TObject);
var
   F                    : TFMul;
   LgFiche              : integer;

begin
  F                     := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
       messagealerte('Aucun élément sélectionné');
       exit;
  end;

  if F.FListe.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not TFMul(Ecran).FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements');
  {$ENDIF}
     F.Q.First;
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.Q.recordcount,True,True);
     While not F.Q.Eof do
     begin
        CreationVisa (CodeCiblage, F.Q.FindField('CCM_CODE').AsString, F.Q.FindField('CCM_NATURE').AsString, '002', F.Q.FindField('CCM_NUMERO').AsString);
        F.Q.Next;
        MoveCurProgressForm('');
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.FListe.nbSelected-1,True,True);
     for LgFiche := 0 to F.FListe.nbSelected-1 do
     begin
        F.FListe.GotoLeBookmark(LgFiche);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}
        CreationVisa (CodeCiblage, F.Q.FindField('CCM_CODE').AsString, F.Q.FindField('CCM_NATURE').AsString, '002', F.Q.FindField('CCM_NUMERO').AsString);
        MoveCurProgressForm('');
     end;
  end;


  FiniMoveProgressForm;
  //F.Q.UpdateCriteres;
  F.BChercheClick(Ecran);

  F.FListe.ClearSelected;

end;




procedure TOF_UTOFRTTOPCONTACT.BenleveOnclick(Sender: TObject);
var
   F                    : TFMul;
   LgFiche              : integer;


begin

  F                     := TFMul(Ecran);
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
     F.Q.First;
     InitMoveProgressForm(nil,'Traitement en cours ...','',F.Q.recordcount,True,True);
     While not F.Q.Eof do
     begin
        if F.Q.FindField('CCM_NATURE').AsString='SUS' then
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('CCM_CODE').AsString+'"')
        else
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_CONTACT="'+F.Q.FindField('CCM_CODE').AsString+
                      '" AND RVB_NUMEROCONTACT="'+F.Q.FindField('CCM_NUMERO').AsString+'"');

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
        if F.Q.FindField('CCM_NATURE').AsString='SUS' then
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_SUSPECT="'+F.Q.FindField('CCM_CODE').AsString+'"')
        else
           ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                      '" AND RVB_CONTACT="'+F.Q.FindField('CCM_CODE').AsString+
                      '" AND RVB_NUMEROCONTACT="'+F.Q.FindField('CCM_NUMERO').AsString+'"');
        MoveCurProgressForm('');
     end;
  end;


  FiniMoveProgressForm;
  //F.Q.UpdateCriteres;
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;


end;


procedure TOF_UTOFRTTOPCONTACT.LaListeOnDblClick(Sender: TObject);
begin
  if VarIsNull(GetField('CCM_CODE')) then
     exit;

  if GetField('CCM_CODECIBLAGE') <> '' then
  begin
     if GetField('CCM_NATURE') = 'SUS' then
        ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                   '" AND RVB_SUSPECT="'+GetField('CCM_CODE')+'"')
     else
        ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                   '" AND RVB_CONTACT="'+GetField('CCM_CODE')+'" AND RVB_NUMEROCONTACT="'+IntToStr(GetField('CCM_NUMERO'))+'"')
  end
  else
     CreationVisa (CodeCiblage, GetField('CCM_CODE'), GetField('CCM_NATURE'), '002', GetField('CCM_NUMERO'));

  TFMul(ecran).BChercheClick(Ecran);



end;


Initialization
  registerclasses ( [ TOF_UTOFRTTOPCONTACT ] ) ;
end.
