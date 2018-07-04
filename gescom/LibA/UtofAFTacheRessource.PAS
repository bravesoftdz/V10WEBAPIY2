{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFTACHERESSOURCE ()
Mots clefs ... : TOF;AFTACHERESSOURCE
*****************************************************************}
Unit UtofAfTACHERESSOURCE;

Interface
                          
Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UTOb,
     AffaireUtil, CalcOleGenericAff,UtilTaches;

Type
  TOF_AFTACHERESSOURCE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private
      fGD_Ressource   : THGrid;
      fTobRessource   : Tob;
      fStAffaire      : String;
      fStNumeroTache  : String;
      fStTiers        : String;
      procedure Load;
      procedure LoadRessource;
      procedure LoadQteRealiseAct;
      procedure LoadQtePla;
      procedure FormatGrid;
      procedure TraitementArgument(pStArgument : String);
  end ;

Function AFLanceFicheAFTacheRessource(Argument : String) : String;

Implementation

Function AFLanceFicheAFTacheRessource(Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFTACHERESSOURCE','', '',Argument);
end;

procedure TOF_AFTACHERESSOURCE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFTACHERESSOURCE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFTACHERESSOURCE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFTACHERESSOURCE.OnLoad ;
begin
  Inherited;
  Load;                                                      
  FormatGrid;
  fTobRessource.PutGridDetail(fGD_Ressource, True, True, fGD_Ressource.Titres[0], True)
end ;

procedure TOF_AFTACHERESSOURCE.OnArgument (S : String ) ;
var
  VStAffFormat: String;
  vStaffiche  : String;
  vStAffLib   : String;

begin
  Inherited ;

  TraitementArgument(S);
  VStAffFormat := CodeAffaireAffiche(fStAffaire,' ');
                                       
  vStaffiche := format('%s / %s   %s',[fStTiers, VStAffFormat, vStAffLib]);
  SetcontrolText('TLIBCLIAFF', vStaffiche);

end;

procedure TOF_AFTACHERESSOURCE.OnClose ;
begin
  Inherited;
end;

procedure TOF_AFTACHERESSOURCE.Load;
begin
  LoadRessource;
  LoadQteRealiseAct;
  LoadQtePla;
end;

procedure TOF_AFTACHERESSOURCE.LoadRessource;
var
  vSt : String;
  vQr : TQuery;

begin

  vSt := 'SELECT ATR_AFFAIRE, ATA_ARTICLE AS ATR_ARTICLE, ATR_RESSOURCE, ';
  vSt := vSt + ' ARS_LIBELLE, ATR_QTEINITIALE, ATR_QTEAPLANIFIER, ';
  vSt := vSt + ' 0.00 AS QTEAFFECTE, 0.00 AS QTEREALISEPLA, 0.00 AS QTEREALISEACT';
  vSt := vSt + ' FROM TACHE, TACHERESSOURCE, RESSOURCE';
  vSt := vSt + ' WHERE ATR_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND ATR_NUMEROTACHE = ' + fStNumeroTache;
  vSt := vSt + ' AND ATR_RESSOURCE = ARS_RESSOURCE';
  vSt := vSt + ' AND ATA_AFFAIRE = ATR_AFFAIRE';
  vSt := vSt + ' AND ATA_NUMEROTACHE = ATR_NUMEROTACHE';

  fTobRessource := TOB.Create('Tob Viewer Mother', Nil, -1);
  vQr := nil;
  try
    vQr := OpenSQL(vSt, True);
    if Not vQr.Eof then
      fTobRessource.LoadDetailDB('mes ressources', '', '', vQr, False, True);
  finally
    Ferme(vQr);
  end;
end;

procedure TOF_AFTACHERESSOURCE.LoadQteRealiseAct;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobAct   : Tob;
  i         : Integer;

Begin

  vSt := vSt + ' SELECT ACT_AFFAIRE, ACT_ARTICLE, ACT_RESSOURCE AS ATR_RESSOURCE, ARS_LIBELLE, ';
  vSt := vSt + ' 0.00 AS ATR_QTEINITIALE, 0.00 ATR_QTEAPLANIFIER, ';
  vSt := vSt + ' 0.00 AS QTEAFFECTE, 0.00 AS QTEREALISEPLA, SUM(ACT_QTE) AS QTEREALISEACT';
  vSt := vSt + ' FROM ACTIVITE, RESSOURCE';
  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS <> "A"';
  vSt := vSt + ' AND AFF_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_RESSOURCE = ARS_RESSOURCE ';
  vSt := vSt + ' AND ACT_FONCTIONRES = AFO_FONCTION ';
  vSt := vSt + ' AND ACT_TYPEARTICLE = "PRE" ';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_RESSOURCE, ARS_LIBELLE';

  vTobAct := TOB.create('Tob Viewer Mother', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAct.LoadDetailDB('mes ressources','','',vQR,False,True);

    for i := vTobAct.Detail.count -1 downto 0 do
      begin
        vTob := fTobRessource.FindFirst(['ATR_AFFAIRE' , 'ATR_ARTICLE', 'ATR_RESSOURCE'],
                                        [vTobAct.Detail[i].GetValue('ACT_AFFAIRE'),
                                         vTobAct.Detail[i].GetValue('ACT_ARTICLE'),
                                         vTobAct.Detail[i].GetValue('ACT_RESSOURCE')], true);

        if vTob <> nil then
          vTob.PutValue('QTEREALISEACT', vTobAct.Detail[i].GetValue('QTEREALISEACT'))

        // ajouter l'enregistrement dans la tob !
        else
          vTobAct.Detail[i].ChangeParent(fTobRessource, -1);

      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
    vTobAct := nil;
  End;
end;

procedure TOF_AFTACHERESSOURCE.LoadQtePla;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobRes   : Tob;
  i         : Integer;

Begin

  vSt := vSt + ' SELECT APL_AFFAIRE, APL_RESSOURCE, ';
  vSt := vSt + ' SUM(APL_QTEREALISE) AS QTEREALISEPLA,';
  vSt := vSt + ' SUM(APL_QTEPLANIFIEE) AS QTEAFFECTE,';
  vSt := vSt + ' FROM AFPLANNING';
  vSt := vSt + ' WHERE APL_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND APL_NUMEROTACHE = ' + fStNumeroTache;
  vSt := vSt + ' GROUP BY APL_AFFAIRE, APL_RESSOURCE';

  vTobRes := TOB.create('Tob Viewer Mother', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);            
    vTobRes.LoadDetailDB('mes ressources','','',vQR,False,True);

    for i := vTobRes.Detail.count -1 downto 0 do
      begin
        vTob := fTobRessource.FindFirst(['ATR_AFFAIRE' , 'ATR_RESSOURCE'],
                                        [vTobRes.Detail[i].GetValue('APL_AFFAIRE'),
                                         vTobRes.Detail[i].GetValue('APL_RESSOURCE')], true);

        if vTob <> nil then
          begin
            vTob.PutValue('QTEREALISEPLA', vTobRes.Detail[i].GetValue('QTEREALISEPLA'));
            vTob.PutValue('QTEAFFECTE', vTobRes.Detail[i].GetValue('QTEAFFECTE'));
          end
        // ajouter l'enregistrement dans la tob !
        else
          vTobRes.Detail[i].ChangeParent(fTobRessource, -1);

      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobRes.Free;
    vTobRes := nil;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : formatage de la grille du plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_AFTACHERESSOURCE.FormatGrid;
var
  i : Integer;
begin

  fGD_Ressource := THGrid(getcontrol('GD_Ressource'));

  fGD_Ressource.rowcount := fTobRessource.Detail.count + 1;
  fGD_Ressource.colcount := 7;

  fGD_Ressource.ColWidths[0] := 75;
  fGD_Ressource.ColWidths[1] := 75;
  fGD_Ressource.ColWidths[2] := 75;
  fGD_Ressource.ColWidths[3] := 75;
  fGD_Ressource.ColWidths[4] := 75;
  fGD_Ressource.ColWidths[5] := 75;
  fGD_Ressource.ColWidths[6] := 75;

  fGD_Ressource.CellValues[0,0] := 'Code Res.';
  fGD_Ressource.CellValues[1,0] := 'Libellé Res.';
  fGD_Ressource.CellValues[2,0] := 'Prévu';
  fGD_Ressource.CellValues[3,0] := 'Planifié';
  fGD_Ressource.CellValues[4,0] := 'Réalisé Planning';
  fGD_Ressource.CellValues[5,0] := 'Réalisé Activité';
  fGD_Ressource.CellValues[6,0] := 'RAP';

  // lignes
  for i := 1 to fGD_Ressource.rowCount - 1 do
    fGD_Ressource.RowHeights[i] := 17;

  fGD_Ressource.RowHeights[0] := 22;

  // alignement
  for i := 2 to fGD_Ressource.ColCount -1 do
    begin
      fGD_Ressource.ColAligns[i] := TaRightJustify;
      fGD_Ressource.ColTypes[i] := 'R';
    end;

  fGD_Ressource.ColEditables[0] := False;
  fGD_Ressource.ColEditables[1] := False;
  fGD_Ressource.ColEditables[3] := False;
  fGD_Ressource.ColEditables[4] := False;
  fGD_Ressource.ColEditables[5] := False;

  fGD_Ressource.ColEditables[2] := true;
  fGD_Ressource.ColEditables[6] := true;

end;

procedure TOF_AFTACHERESSOURCE.TraitementArgument(pStArgument : String);
var
  Tmp     : String;
  champ   : String;
  valeur  : String;

begin
  // traitement des arguments
  Tmp:=(Trim(ReadTokenSt(pStArgument)));
  While (Tmp <>'') do
    Begin
      If Tmp<>'' then
        Begin
          DecodeArgument(Tmp, Champ, valeur);
          If Champ='AFFAIRE'      then fStAffaire := valeur else
          If Champ='NUMEROTACHE'  then fStNumeroTache := valeur else
          If Champ='TIERS'  then fStTiers := valeur;
        End;                                        
        Tmp:=(Trim(ReadTokenSt(pStArgument)));
    End;
end;

Initialization
  registerclasses ( [ TOF_AFTACHERESSOURCE ] ) ;
end.
