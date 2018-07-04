{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : STDALIGNEPLAN ()
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... : TOF;STDALIGNEPLAN
*****************************************************************}
Unit uTofStdAlignePlan ;

Interface

Uses Windows,
     StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     uLanceProcess, // LanceProcessServer
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     HSysMenu,
     HTB97,
     UTOF,
     uTOB,
     ParamSoc,       // GetParamSoc
{$IFDEF COMPTA}
     AligneStd,      // TAligneStd
{$ENDIF}
     Ent1;

const
  COL_PLAN        = 1;
  LIG_GENERAUX    = 1;
  LIG_TIERS       = 2;
  LIG_JOURNAUX    = 3;
  LIG_GUIDES      = 4;
  LIG_SECTIONS    = 5;
  LIG_STDEDITION  = 6;
  LIG_RUPTEDITION = 7;
  LIG_CPTECORRESP = 8;
  LIG_TABLIBRES   = 9;
  LIG_TVA         = 10;
  LIG_LIBELLEAUTO = 11;


Type
  TOF_STDALIGNEPLAN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      FGrille          : THGrid;
      FGrilleResultat  : THGrid;
      FHSystemMenu     : THSystemMenu;
      FResultat        : TOB;
      FCBRazStdEdition : TCheckBox;
      FBoMAJ           : Boolean;

      procedure   AfficheCommentaire ( St : string );
{$IFDEF COMPTA}
      procedure   ChargeAAligner ( Alignement : TAligneStd );
{$ENDIF}

      procedure   OnInfoAlignement (Sender : TObject; Err : integer; Msg : string);
      procedure   OnDblClickResultat (Sender : TObject);
      procedure   InitGrilleSaisie;
      procedure   OnFormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure   OnClickMemePlan (Sender : TObject );
      procedure   AffecteMemePlan;
      procedure   OnRowEnterFGrille(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure   RemplitTobALigneStd( var vTobAligneStd : Tob );

  end ;

{$IFDEF COMPTA}
function  CPAlignementMultiDossier( vStCheminTobAligneStd : string ) : Boolean;
{$ENDIF}
function CPLanceFiche_AlignementSurStandard( vBoMaj : Boolean = False ) : String;
procedure MiseAJourPlanRefence;

Implementation

uses
{$IFDEF COMPTA}
     CPGeneraux_TOM, // FicheGene
     CPTiers_TOM,    // FicheTiers
     CPJournal_TOM,
     CPGUIDE_TOM,
     CPSection_TOM,
{$ENDIF}
{$IFDEF MODENT1}
      CPProcGen,
{$ENDIF MODENT1}
     uLibStdCpta,    // InitTonParamProcessServer
     uLibWindows;    // IF


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
Function CPAlignementMultiDossier( vStCheminTobAligneStd : string ) : Boolean;
var lTobAligneStd : Tob;
    lAligneStd : TAligneStd;
begin
  try
    Result := False;
    lAligneStd := TAligneStd.Create;
    lAligneStd.FRazFiltres := False;
    lAligneStd.OkProgressForm := True;
    lAligneStd.OnInformation := nil;

    lTobAligneStd := Tob.Create('', nil, -1);
    TobLoadFromFile( vStCheminTobAligneStd, nil, lTobAligneStd );

    if lTobAligneStd.Detail.Count > 0 then
    begin
      // GCO - 17/03/2006 - FQ 17635
      if lTobAligneStd.FieldExists('RAZFILTRES') then
        lAligneStd.FRazFiltres := lTobAligneStd.GetBoolean('RAZFILTRES');

      lAligneStd.ChargeAligneStdAvecTob( lTobAligneStd );
      lAligneStd.Execute;

      if lAligneStd.LastError >= 0 then
        Result := True;
    end;
  finally
    FreeAndNil( lTobAligneStd );
    FreeAndNil( lAligneStd );

    // Interprétation du Resultat
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_AlignementSurStandard( vBoMaj : Boolean = False ) : String;
begin
  if vBoMaj then
    Result := AGLLanceFiche ('CP','STDALIGNEPLAN','','', 'OUI')
  else
    Result := AGLLanceFiche ('CP','STDALIGNEPLAN','','', 'NON');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure MiseAJourPlanRefence;
var lVarTemp : Variant;
begin
  // Demande d'alignement sur le standard de référence
  if (GetParamSocSecur('SO_CPOKMAJPLAN', False)) and (GetParamSocSecur('SO_NUMPLANREF', 0) > 0) then
  begin
    lVarTemp := GetColonneSQl('STDCPTA', 'STC_DATEMODIF', 'STC_NUMPLAN = ' + IntToStr(GetParamSoc('SO_NUMPLANREF')));

    if IsValidDateHeure( lVarTemp ) then
    begin
      if StrToDateTime(GetParamSoc('SO_CPDATEDERNMAJPLAN')) < StrToDateTime(lVarTemp) then
      begin
        if PgiAsk('Le standard de référence du dossier à été modifié.' + #13 + #10 +
                  'Voulez vous effectuer un alignement ?', 'Alignement des standards') = MrYes then
        begin
          V_PGI.ZoomOLE := TRUE;
          CPLanceFiche_AlignementSurStandard( True );
          V_PGI.ZoomOLE := False;
        end
        else
          SetParamSoc('SO_CPDATEDERNMAJPLAN', Now);
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_STDALIGNEPLAN.OnArgument (S : String ) ;
var lStArgument : string;
begin
  Inherited ;
  lStArgument := ReadTokenSt(S);

  FBoMAJ := IIF( lStArgument = 'OUI', True, False);

  FHSystemMenu := THSystemMenu (TFVierge(ECRAN).HMTrad);
  ECRAN.OnKeyDown := OnFormKeyDown;
  TToolBarButton97(GetControl('BMEMEPLAN')).OnClick := OnClickMemePlan;
  FGrille := THGrid (GetControl ('FGRILLE'));
  InitGrilleSaisie;
  FGrilleResultat := THGrid (GetControl ('FRESULTAT'));
  FGrilleResultat.ColWidths[0]:=20;
  FGrilleResultat.OnDblClick := OnDblClickResultat;
  FHSystemMenu.ResizeGridColumns ( FGrilleResultat );
  FResultat := TOB.Create ('', nil, -1);
  FCBRazStdEdition := TCheckBox(GetControl('CBRAZSTDEDITION', True));
  FGrille.OnRowEnter := OnRowEnterFGrille;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_STDALIGNEPLAN.OnLoad ;
begin
  SetControlProperty('PRESULTAT','TabVisible',False);
  AfficheCommentaire('');
  Inherited ;
end ;


procedure TOF_STDALIGNEPLAN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_STDALIGNEPLAN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_STDALIGNEPLAN.OnUpdate ;
var lNumStd    : integer;
{$IFDEF DP}
  lTobParam : Tob;
  lsTmpPath : String;
{$ELSE}
{$IFDEF EAGLCLIENT}
  lTobResult : Tob;
  lTobParam  : Tob;
{$ELSE}
  Alignement : TAligneStd;
{$ENDIF}
{$ENDIF}

begin
  SetControlProperty('PRESULTAT','TabVisible',False);
  FGrille.ValCombo.Hide;

  if (FGrille.CellValues[COL_PLAN,LIG_GUIDES]<>'') and (FGrille.CellValues[COL_PLAN,LIG_JOURNAUX]<>FGrille.CellValues[COL_PLAN,LIG_GUIDES]) then
  begin
    PGIBox ('Alignement impossible.#10#13Les guides doivent être alignés avec les journaux correspondants.', ECRAN.Caption);
    Exit;
  end;

{$IFDEF DP}

{$ELSE}
  // GCO - 06/09/2006 - FQ 18674
  if (not GetParamSocSecur('SO_CPPCLSANSANA',false)) then
  begin // Gestion de l'analytique
    if (FGrille.CellValues[COL_PLAN,LIG_GENERAUX] <> '') and
       (FGrille.CellValues[COL_PLAN,LIG_SECTIONS] <> FGrille.CellValues[COL_PLAN,LIG_GENERAUX]) then
    begin
      PGIBox ('Alignement impossible.#10#13Les sections doivent être alignées avec les comptes généraux.', ECRAN.Caption);
      Exit;
    end;
  end
  else
  begin
    // GCO - 12/07/2006 - FQ 15744
    // Le Dossier ne gère pas l'analytique et l'utilisateur veut aligner les sections
    if (FGrille.CellValues[COL_PLAN,LIG_SECTIONS] <> '') then
    begin
      PGIBox ('Alignement impossible.#10#13L''analytique n''étant pas gérée dans ce dossier, ' +
              'l''alignement des sections n''est pas possible.', ECRAN.Caption);
      FGrille.CellValues[COL_PLAN,LIG_SECTIONS] := '';
      Exit;
    end;
  end;
{$ENDIF}

  // Contrôle que les filtres du standard choisi sont aux formats XML
  if (FGrille.CellValues[COL_PLAN, LIG_STDEDITION] <> '') then
  begin
  //{$IFDEF EAGLCLIENT}
  //  PgiInfo('L''alignement des standards d''éditions n''est pas disponible pour le moment.' + #13 + #10 +
//            'Ce traitement ne sera pas effectué.' , 'Alignement des standards');
  //  Exit;
 // {$ELSE}
    lNumStd := StrToInt(FGrille.CellValues[COL_PLAN, LIG_STDEDITION]);
    // GCO - 05/10/2004 - FQ 14330
    if PresenceAncienFiltreDansSTD( lNumStd ) then
    begin
      PgiInfo(MessageAncienFiltreDansSTD( lNumStd, 2), Ecran.Caption);
      Exit;
    end
  //{$ENDIF EAGLCLIENT}
  end;

{$IFDEF DP}
  lTobParam := Tob.Create('', nil, -1);

  // GCO - 17/03/2006 - FQ 17635
  if FCBRazStdEdition.Checked then
    lTobParam.AddChampSupValeur('RAZFILTRES', 'X')
  else
    lTobParam.AddChampSupValeur('RAZFILTRES', '-');

  lsTmpPath := TempFileName;
  RemplitTobALigneStd( lTobParam );
  lTobParam.SaveToFile( lsTmpPath, False, True, True);
  TFVierge(Ecran).Retour := 'OUI; /ALIGNESTD=' + lsTmpPath;
  Ecran.Close;
{$ELSE}

{$IFDEF EAGLCLIENT}
  lTobParam := InitTobParamProcessServer;

  // GCO - 31/08/2007 - FQ 21289
  if FCBRazStdEdition.Checked then
    lTobParam.AddChampSupValeur('RAZFILTRES', 'X')
  else
    lTobParam.AddChampSupValeur('RAZFILTRES', '-');

  // Paramètres pour l'appel du process server :
  RemplitTobALigneStd( lTobParam );

  lTobResult := LanceProcessServer('cgiStdCpta', 'ALIGNESTD', 'PASDEPARAM', lTobParam, True ) ;
  if lTobResult = nil then
  begin
    PgiError('Erreur dans les informations de retour...', 'ProcesseServer : cgiStdCpta');
  end
  else
  begin
    if lTobResult.GetValue('RESULT') = 'PASOKALIGNESTD' then
      PgiInfo('Traitement annulé. Erreur pendant l''alignement.', Ecran.Caption)
    else
      PgiInfo('Alignement terminé avec succès.', Ecran.Caption);

    lTobResult.PutGridDetail(FGrilleResultat,False,False,';Nature;Valeur');
    SetControlProperty('PRESULTAT','TabVisible',(lTobResult.Detail.Count > 0));
    FreeAndNil(lTobResult);
  end;
{$ELSE}
  // Alignement
  Alignement := TAligneStd.Create;
  Alignement.FRazFiltres := FCBRazStdEdition.Checked;
  Alignement.OnInformation := OnInfoAlignement;
  ChargeAAligner ( Alignement );
  Alignement.Execute;
  Alignement.ChargeResultat (FResultat);

  if Alignement.LastError = 0 then
  begin
    AvertirMultiTable('TTJOURNAL'); // GCO - 20/07/2005 - FQ 15424
    PGIInfo('Alignement terminé avec succès.',Ecran.Caption);
  end
  else
    PGIBox(Alignement.LastErrorMsg,Ecran.Caption);

  FResultat.PutGridDetail(FGrilleResultat, False, False, ';Nature;Valeur');
  SetControlProperty('PRESULTAT', 'TabVisible', (FResultat.Detail.Count > 0));
  Alignement.Free;
{$ENDIF EAGLCLIENT}
  AfficheCommentaire('');
{$ENDIF DP}
end ;

procedure TOF_STDALIGNEPLAN.AfficheCommentaire(St: string);
begin
  SetControlProperty('FCOMMENTAIRE','Caption',St);
  Application.ProcessMessages;
end;

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
procedure TOF_STDALIGNEPLAN.ChargeAAligner ( Alignement : TAligneStd );
begin
  Alignement.AjouteTAF ( alGeneraux,    FGrille.CellValues[COL_PLAN,LIG_GENERAUX]);
  Alignement.AjouteTAF ( alTiers,       FGrille.CellValues[COL_PLAN,LIG_TIERS]);
  Alignement.AjouteTAF ( alJournaux,    FGrille.CellValues[COL_PLAN,LIG_JOURNAUX]);
  Alignement.AjouteTAF ( alGuides,      FGrille.CellValues[COL_PLAN,LIG_GUIDES]);
  Alignement.AjouteTAF ( alSections,    FGrille.CellValues[COL_PLAN,LIG_SECTIONS]);
  Alignement.AjouteTAF ( alStdEdition,  FGrille.CellValues[COL_PLAN,LIG_STDEDITION]);
  Alignement.AjouteTAF ( alRuptEdition, FGrille.CellValues[COL_PLAN,LIG_RUPTEDITION]);
  Alignement.AjouteTAF ( alCpteCorresp, FGrille.CellValues[COL_PLAN,LIG_CPTECORRESP]);
  Alignement.AjouteTAF ( alTabLibres,   FGrille.CellValues[COL_PLAN,LIG_TABLIBRES]);
  // GCO - 09/05/2005
  Alignement.AjouteTAF ( alTva,         FGrille.CellValues[COL_PLAN, LIG_TVA]);
  Alignement.AjouteTAF ( alLibelleAuto, FGrille.CellValues[COL_PLAN, LIG_LIBELLEAUTO]);
  // FIN GCO
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_STDALIGNEPLAN.RemplitTobALigneStd( var vTobAligneStd : Tob );
var lTobFille : Tob;
begin
  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_GENERAUX');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_GENERAUX]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_TIERS');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_TIERS]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_JOURNAUX');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_JOURNAUX]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_GUIDES');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_GUIDES]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_SECTIONS');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_SECTIONS]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_STDEDITION');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_STDEDITION]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_RUPTEDITION');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_RUPTEDITION]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_CPTECORRESP');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_CPTECORRESP]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_TABLIBRES');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_TABLIBRES]);

  // GCO - 05/05/2005 - Ajout de la TVA et des libellés AUTO
  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_TVA');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_TVA]);

  lTobFille := Tob.Create( '', vTobAligneStd, -1);
  lTobFille.AddchampSupValeur( 'TRAITEMENT', 'LIG_LIBELLEAUTO');
  lTobFille.AddchampSupValeur( 'NUMSTD', FGrille.CellValues[COL_PLAN,LIG_LIBELLEAUTO]);
  // FIN GCO
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_STDALIGNEPLAN.OnInfoAlignement(Sender: TObject; Err: integer; Msg: string);
begin
  SetControlProperty('FCOMMENTAIRE','Caption',Msg);
  Application.ProcessMessages;
end;

procedure TOF_STDALIGNEPLAN.OnDblClickResultat(Sender: TObject);
{$IFDEF COMPTA}
var T : TOB;
{$ENDIF}
begin
{$IFDEF COMPTA}
  T := TOB (FGrilleResultat.Objects[0,FGrilleResultat.Row]);
  if T.GetValue('Table')='GENERAUX' then
    FicheGene(nil,'',T.GetValue('Valeur'),taConsult,0)
  else if T.GetValue('Table')='TIERS' then
    FicheTiers(nil,'',T.GetValue('Valeur'),taConsult,0)
  else if T.GetValue('Table')='JOURNAL' then
    FicheJournal(nil,'',T.GetValue('Valeur'),taConsult,0)
  else if T.GetValue('Table')='GUIDE' then
    ParamGuide ( T.GetValue('Valeur'),'NOR',taConsult )
  else if T.GetValue('Table')='SECTION' then
    FicheSection(nil,'', T.GetValue('Valeur'), taConsult, 0);
{$ENDIF}
end;

procedure TOF_STDALIGNEPLAN.InitGrilleSaisie;
var lStNumPlanRef : string;
    lStCellValue  : string;
begin
  FGrille.RowCount := 12;
  FGrille.ColWidths[0]:=100;
  FGrille.Cells[0,1]  := TraduireMemoire('Généraux');
  FGrille.Cells[0,2]  := TraduireMemoire('Auxiliaires');
  FGrille.Cells[0,3]  := TraduireMemoire('Journaux');
  FGrille.Cells[0,4]  := TraduireMemoire('Guides');
  FGrille.Cells[0,5]  := TraduireMemoire('Sections');
  FGrille.Cells[0,6]  := TraduireMemoire('Standards d''éditions');
  FGrille.Cells[0,7]  := TraduireMemoire('Plans de rupture des éditions');
  FGrille.Cells[0,8]  := TraduireMemoire('Comptes de correspondance');
  FGrille.Cells[0,9]  := TraduireMemoire('Tables libres');
  FGrille.Cells[0,10] := TraduireMemoire('TVA');
  FGrille.Cells[0,11] := TraduireMemoire('Libellés automatiques');
  FGrille.ColFormats[1] := 'CB=CPSTDCPTA||Aucun';

  lStNumPlanRef := GetParamSocSecur('SO_NUMPLANREF',''); {Lek 100206}

  if ( StrToInt(lStNumPlanRef) > 0) and FBoMAJ then
    lStCellValue := lStNumPlanRef
  else
    lStCellValue := '';

  FGrille.CellValues[1,1]  := lStCellValue;
  FGrille.CellValues[1,2]  := lStCellValue;
  FGrille.CellValues[1,3]  := lStCellValue;
  FGrille.CellValues[1,4]  := lStCellValue;
  FGrille.CellValues[1,5]  := lStCellValue;
  FGrille.CellValues[1,6]  := lStCellValue;
  FGrille.CellValues[1,7]  := lStCellValue;
  FGrille.CellValues[1,8]  := lStCellValue;
  FGrille.CellValues[1,9]  := lStCellValue;
  FGrille.CellValues[1,10] := lStCellValue;
  FGrille.CellValues[1,11] := lStCellValue;

  FHSystemMenu.ResizeGridColumns ( FGrille );
end;


procedure TOF_STDALIGNEPLAN.OnFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ( csDestroying in Ecran.ComponentState ) then Exit ;
  case Key of
    VK_F10 :
      begin
        TToolBarButton97(GetControl('BVALIDER')).Click;
        Key := 0;
      end;
    Ord('A')  :
      if Shift = [ssCtrl] then
      begin
        AffecteMemePlan;
        Key := 0;
      end;
  end;
end;

procedure TOF_STDALIGNEPLAN.AffecteMemePlan;
var i, lDebut : integer;
begin
  FGrille.ValCombo.Hide;
  lDebut := FGrille.Row;
  for i := lDebut+1 to FGrille.RowCount - 1 do
    FGrille.CellValues[COL_PLAN,i] := FGrille.CellValues[COL_PLAN,lDebut];
end;

procedure TOF_STDALIGNEPLAN.OnClickMemePlan(Sender: TObject);
begin
  AffecteMemePlan;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_STDALIGNEPLAN.OnRowEnterFGrille(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  (*
  if Ou = 5 then
  begin
    FGrille.ColFormats[1] := '';
  end
  else
  begin
    FGrille.ColFormats[1] := 'CB=CPSTDCPTA||Aucun';
  end;*)
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_STDALIGNEPLAN.OnClose ;
begin
  FResultat.Free;
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_STDALIGNEPLAN ] ) ;
end.

