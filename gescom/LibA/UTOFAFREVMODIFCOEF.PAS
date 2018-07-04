{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVMODIFCOEF ()
Mots clefs ... : TOF;AFREVMODIFCOEF
*****************************************************************}
Unit uTofAfRevModifCoef ;
 
Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,mul,
{$Else}
     MainEagl,eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,GCMZSUtil,HTB97,UTofAfBaseCodeAffaire,utofafrevision,
     EntGC, UtilMulTrt, UTob, UAFO_REVPRIXCALCULCOEF, utilRevision,
     Dicobtp,  Ed_tools;

Type

  TOF_AFREVMODIFCOEF =  Class (TOF_AFBASECODEAFFAIRE)
  private
    fBoAppliquer          : Boolean;
    procedure ModifCoefEnSerie(Sender: TObject);
    procedure TraitementCoefEnSerie(Sender: TObject);

    procedure RB_APPLIQUEROnClick(Sender: TObject);
    procedure RB_APPLIQUERDERNIEROnClick(Sender: TObject);
    procedure RB_ANNULEROnClick(Sender: TObject);

  public                      
    LaListe : THGrid ;
    procedure LaListeDblClick(sender : Tobject) ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  end ;

procedure AFLanceFicheModifCoef(s : String);
 
Implementation

procedure TOF_AFREVMODIFCOEF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('GL_TIERS'));
end;

procedure TOF_AFREVMODIFCOEF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVMODIFCOEF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVMODIFCOEF.OnUpdate ;
begin
  Inherited ;
end ;
 
procedure TOF_AFREVMODIFCOEF.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_AFREVMODIFCOEF.LaListeDblClick(sender : Tobject) ;
Var s  : string;
begin

  s := ' WHERE AFR_AFFAIRE ="' + GetField('AFF_AFFAIRE') + '"';
  s := s + ' AND AFR_NUMEROLIGNE = "' + inttostr(GetField('AFR_NUMEROLIGNE')) + '"';
  s := s + ' AND AFR_FORCODE = "' + GetField('AFR_FORCODE') + '"';

  if GetField('GL_NUMERO') = 0 then
    AglLanceFicheAFREVISION('','Clause_where=' + s + ';AFFAIRE;ACTION=MODIFICATION')
  else
    AglLanceFicheAFREVISION('','Clause_where=' + s + ';LIGNEAFFAIRE;ACTION=MODIFICATION') ;
  TFMul(Ecran).ChercheClick;
end;

procedure TOF_AFREVMODIFCOEF.ModifCoefEnSerie(Sender: TObject);
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin MessageAlerte('Aucun élément sélectionné'); exit; end;

  TheModifLot := TO_ModifParLot.Create;
  TheModifLot.F := F.FListe;
  TheModifLot.Q := F.Q;
  TheModifLot.NatureTiers := '';
  TheModifLot.Nature := 'AFF';
  TheModifLot.Titre := Ecran.Caption;
  TheModifLot.TableName:='AFREVISION';
  TheModifLot.FCode := 'AFR_AFFAIRE;AFR_FORCODE;AFR_NUMEROLIGNE' ;

  V_PGI.ExtendedFieldSelection:='1';
  ModifieEnSerie(TheModifLot, Parametrages) ;
  if F.FListe.AllSelected then F.FListe.AllSelected:=False
                          else F.FListe.ClearSelected;
  F.bSelectAll.Down := False ;

  // relancer le load pour raffraichir le mul
  TFMul(Ecran).BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2003
Modifié le ... :   /  /
Description .. : application en serie des coefficients
Mots clefs ... :
*****************************************************************}
procedure TOF_AFREVMODIFCOEF.TraitementCoefEnSerie(Sender: TObject);
Var
  F         : TFMul ;
  ret       : T_ErrTraiteEnreg;
  vTob      : Tob;
  i         : Integer;
  vRev      : TCALCULCOEF;
  vBoResult : Boolean;
  vSLLog    : TStringList;
  vStRes    : String;

begin

  vBoresult := True;
  vTob := Tob.Create('AFREVISION',NIL,-1);
  try
    F := TFMul(Ecran);
    if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

    if fBoAppliquer then
      vStRes := agllancefiche('AFF','AFREVTRAITCOEF','','','');

    if vStRes = '3' then exit;

    ret := TraiteEnregMulListe (TFMul(Ecran), 'AFR_AFFAIRE','AFREVMODIFCOEF', vTob, True);
    If ret <> tteOK then exit ;

    InitMoveProgressForm(nil, 'Traitement en cours...', '', vTob.detail.count, false, true);
    Try
      for i := 0 to vTob.Detail.count -1 do
      begin
        MoveCurProgressForm('');
        vRev := TCALCULCOEF.Create;
        try
          if not vRev.LoadFormule(vTob.detail[i].GetValue('AFR_AFFAIRE'), vTob.detail[i].GetValue('AFR_FORCODE')) then
          begin
            vBoResult := False;
            DebutLog('', vSLLog);
            vSLLog.Add(format(TexteMsg[1], [DateToStr(now),
                                            vTob.detail[i].GetValue('AFR_AFFAIRE'),
                                            vTob.detail[i].GetValue('AFR_FORCODE')]));
            FinLog('', vSLLog);
            break;
          end
          else
          begin
            if fBoAppliquer then
            begin
              case strToInt(vStRes) of
                0 : vRev.AppliquerCoef(false, i = 0); // appliquer
                1 : vRev.AppliquerCoef(true, i = 0); // appliquer le dernier
                2 : if not vRev.DesappliquerCoef(vTob.detail[i].GetValue('AFR_DATECALCCOEF')) then
                      vBoResult := False;
              end;
            end;
          end;

        finally
          vRev.Free;
        end;
      end;
    finally
      FiniMoveProgressForm;
    end;
                  
  if not vBoResult then
  begin
    if strToInt(vStRes) = 2 then
      PGIInfoAF('Un ou plusieurs coefficients n''ont pas pu être appliqués. Vous pouvez consulter le fichier .log.', Ecran.caption)
    else                         
      PGIInfoAF('L''annulation de l''application d''un ou plusieurs coefficients n''a pas pu être effectuée. Vous pouvez consulter le fichier .log.', Ecran.caption);
  end
  else
    begin
      // relancer le load pour raffraichir le mul
      TFMul(Ecran).BCherche.Click;
      PGIInfoAF('Traitement terminé.', Ecran.caption);
    end;

  finally
    vTob.Free;
  end;
end;

procedure TOF_AFREVMODIFCOEF.OnArgument (S : String ) ;
var
  vSt : STring;
begin
  Inherited ;
  fBoAppliquer := S = 'APPLIQUER';
  if fBoAppliquer then
  begin
    TToolBarButton97 (GetControl('BOuvrir')).OnClick := TraitementCoefEnSerie;
    ecran.Caption := traduitGA('Application par lots des coefficients');
    UpdateCaption(Ecran);

    TFMul(Ecran).DBListe := 'AFREVAPPLIQCOEF';
    if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste := 'AFREVAPPLIQCOEF';

    vSt := 'AND ((GL_NATUREPIECEG<>"FAC") AND (GL_NATUREPIECEG<>"FPR")) OR (GL_NATUREPIECEG IS NULL)';
    SetControlText('XX_WHERE', vSt);
  end
  else if S = 'MODIFIER' then
  begin
    TToolBarButton97 (GetControl('BOuvrir')).OnClick  := ModifCoefEnSerie;
    TCheckBox(GetControl('AFR_OKCOEFAPPLIQUE')).checked := false;
    SetControlVisible('AFR_OKCOEFAPPLIQUE', false);
    ecran.Caption := traduitGA('Modification par lots des coefficients');
    UpdateCaption(Ecran);
    vSt := 'AND ((GL_NATUREPIECEG<>"FAC") AND (GL_NATUREPIECEG<>"FPR")) OR (GL_NATUREPIECEG IS NULL)';
    SetControlText('XX_WHERE', vSt);
  end
  else
  begin
    vSt := 'AND ((GL_NATUREPIECEG="FAC") OR (GL_NATUREPIECEG="FPR"))';
    SetControlText('XX_WHERE', vSt);
    TFMul(Ecran).DBListe := 'AFREVCONSULTCOEF';
    if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste := 'AFREVCONSULTCOEF';

    SetControlVisible('BSelectAll', false);
    SetControlVisible('BOuvrir', false);
    SetControlVisible('AFR_OKCOEFAPPLIQUE', false);
    SetControlVisible('RB_APPLIQUER', false);
    SetControlVisible('RB_APPLIQUERDERNIER', false);
    SetControlVisible('RB_ANNULER', false);
    SetControlVisible('PCOMPLEMENT', true);
    SetControlVisible('AFR_COEFREGUL', true);
                                     
    ecran.Caption := traduitGA('Eléments de révision de prix');
    UpdateCaption(Ecran);
    ecran.Refresh;
  end;

  LaListe:=THGrid(GetControl('Fliste')) ;
  LaListe.OnDblClick:= LaListeDblClick;
end;

procedure TOF_AFREVMODIFCOEF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVMODIFCOEF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVMODIFCOEF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVMODIFCOEF.RB_APPLIQUEROnClick(Sender: TObject);
begin

end;

procedure TOF_AFREVMODIFCOEF.RB_APPLIQUERDERNIEROnClick(Sender: TObject);
begin

end;

procedure TOF_AFREVMODIFCOEF.RB_ANNULEROnClick(Sender: TObject);
begin

end;

procedure AFLanceFicheModifCoef(s : String);
begin
  AGLLanceFiche('AFF','AFREVMODIFCOEF','','',s);
end;

Initialization
  registerclasses ( [ TOF_AFREVMODIFCOEF ] ) ;
end.

