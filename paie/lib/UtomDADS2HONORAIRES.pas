{$A8,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q+,R+,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{$WARN SYMBOL_DEPRECATED ON}
{$WARN SYMBOL_LIBRARY ON}
{$WARN SYMBOL_PLATFORM ON}
{$WARN UNIT_LIBRARY ON}
{$WARN UNIT_PLATFORM ON}
{$WARN UNIT_DEPRECATED ON}
{$WARN HRESULT_COMPAT ON}
{$WARN HIDING_MEMBER ON}
{$WARN HIDDEN_VIRTUAL ON}
{$WARN GARBAGE ON}
{$WARN BOUNDS_ERROR ON}
{$WARN ZERO_NIL_COMPAT ON}
{$WARN STRING_CONST_TRUNCED ON}
{$WARN FOR_LOOP_VAR_VARPAR ON}
{$WARN TYPED_CONST_VARPAR ON}
{$WARN ASG_TO_TYPED_CONST ON}
{$WARN CASE_LABEL_RANGE ON}
{$WARN FOR_VARIABLE ON}
{$WARN CONSTRUCTING_ABSTRACT ON}
{$WARN COMPARISON_FALSE ON}
{$WARN COMPARISON_TRUE ON}
{$WARN COMPARING_SIGNED_UNSIGNED ON}
{$WARN COMBINING_SIGNED_UNSIGNED ON}
{$WARN UNSUPPORTED_CONSTRUCT ON}
{$WARN FILE_OPEN ON}
{$WARN FILE_OPEN_UNITSRC ON}
{$WARN BAD_GLOBAL_SYMBOL ON}
{$WARN DUPLICATE_CTOR_DTOR ON}
{$WARN INVALID_DIRECTIVE ON}
{$WARN PACKAGE_NO_LINK ON}
{$WARN PACKAGED_THREADVAR ON}
{$WARN IMPLICIT_IMPORT ON}
{$WARN HPPEMIT_IGNORED ON}
{$WARN NO_RETVAL ON}
{$WARN USE_BEFORE_DEF ON}
{$WARN FOR_LOOP_VAR_UNDEF ON}
{$WARN UNIT_NAME_MISMATCH ON}
{$WARN NO_CFG_FILE_FOUND ON}
{$WARN MESSAGE_DIRECTIVE ON}
{$WARN IMPLICIT_VARIANTS ON}
{$WARN UNICODE_TO_LOCALE ON}
{$WARN LOCALE_TO_UNICODE ON}
{$WARN IMAGEBASE_MULTIPLE ON}
{$WARN SUSPICIOUS_TYPECAST ON}
{$WARN PRIVATE_PROPACCESSOR ON}
{$WARN UNSAFE_TYPE ON}
{$WARN UNSAFE_CODE ON}
{$WARN UNSAFE_CAST OFF}
{$A8,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q+,R+,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{$WARN SYMBOL_DEPRECATED ON}
{$WARN SYMBOL_LIBRARY ON}
{$WARN SYMBOL_PLATFORM ON}
{$WARN UNIT_LIBRARY ON}
{$WARN UNIT_PLATFORM ON}
{$WARN UNIT_DEPRECATED ON}
{$WARN HRESULT_COMPAT ON}
{$WARN HIDING_MEMBER ON}
{$WARN HIDDEN_VIRTUAL ON}
{$WARN GARBAGE ON}
{$WARN BOUNDS_ERROR ON}
{$WARN ZERO_NIL_COMPAT ON}
{$WARN STRING_CONST_TRUNCED ON}
{$WARN FOR_LOOP_VAR_VARPAR ON}
{$WARN TYPED_CONST_VARPAR ON}
{$WARN ASG_TO_TYPED_CONST ON}
{$WARN CASE_LABEL_RANGE ON}
{$WARN FOR_VARIABLE ON}
{$WARN CONSTRUCTING_ABSTRACT ON}
{$WARN COMPARISON_FALSE ON}
{$WARN COMPARISON_TRUE ON}
{$WARN COMPARING_SIGNED_UNSIGNED ON}
{$WARN COMBINING_SIGNED_UNSIGNED ON}
{$WARN UNSUPPORTED_CONSTRUCT ON}
{$WARN FILE_OPEN ON}
{$WARN FILE_OPEN_UNITSRC ON}
{$WARN BAD_GLOBAL_SYMBOL ON}
{$WARN DUPLICATE_CTOR_DTOR ON}
{$WARN INVALID_DIRECTIVE ON}
{$WARN PACKAGE_NO_LINK ON}
{$WARN PACKAGED_THREADVAR ON}
{$WARN IMPLICIT_IMPORT ON}
{$WARN HPPEMIT_IGNORED ON}
{$WARN NO_RETVAL ON}
{$WARN USE_BEFORE_DEF ON}
{$WARN FOR_LOOP_VAR_UNDEF ON}
{$WARN UNIT_NAME_MISMATCH ON}
{$WARN NO_CFG_FILE_FOUND ON}
{$WARN MESSAGE_DIRECTIVE ON}
{$WARN IMPLICIT_VARIANTS ON}
{$WARN UNICODE_TO_LOCALE ON}
{$WARN LOCALE_TO_UNICODE ON}
{$WARN IMAGEBASE_MULTIPLE ON}
{$WARN SUSPICIOUS_TYPECAST ON}
{$WARN PRIVATE_PROPACCESSOR ON}
{$WARN UNSAFE_TYPE ON}
{$WARN UNSAFE_CODE ON}
{$WARN UNSAFE_CAST OFF}
{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DADS2HONORAIRES
Suite ........ : (DADS2HONORAIRES)
Mots clefs ... : TOM;DADS2HONORAIRES
*****************************************************************}
{
PT1-1 : 08/10/2004 VG V_50 Adaptation cahier des charges 2004
PT1-2 : 08/10/2004 VG V_50 Contrôles de cohérence - FQ N°11657
PT2   : 20/10/2004 VG V_50 nom + prénom + raison sociale renseignés
                           simultanément
PT3   : 18/11/2004 VG V_60 Lors de la navigation, les coches n'étaient pas
                           décochées - FQ N°11789
PT4   : 03/01/2005 VG V_60 Contrôles de cohérence - FQ N°11657
PT5   : 07/02/2005 VG V_60 Contrôles de cohérence en validation - FQ N°11971
PT6   : 11/03/2005 VG V_60 Initialisation des champs de l'enregistrement
                           "Honoraires" en création - FQ N°12090
PT7   : 20/10/2005 VG V_60 Suppression du contrôle sur le siret en cas de
                           personne physique - FQ N°12219
PT8   : 04/01/2006 PH V_65 FQ 12766 Erreur création manuelle Honoraires si déjà
                           importés via la compta alors le code honoraire est le
                           compte auxiliaire et l'incrément automatique ne
                           marche pas.
PT9   : 30/03/2006 VG V_65 Correction PT8
PT10  : 11/01/2007 VG V_72 On passe le contrôle du SIRET non bloquant
                           FQ N°13692
PT11  : 13/11/2007 NA V_80 Le ocde section 00 n'existe plus
}
unit UtomDADS2HONORAIRES;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     Fiche,
{$ELSE}
     eFiche,
     MaineAgl,
{$ENDIF}
     UTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
{$IFDEF COMPTA}
     UtilPGI,
     UtilTrans,
{$ELSE}
     PgOutils2,
{$ENDIF}
     PgDADSCommun,
     ParamSoc ;

type
  TOM_DADS2HONORAIRES = class(TOM)
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  private
    AlloForfait, AuAvant, Dispense, Employeur, Logement, Nourriture : TCheckBox;
    NTIC, Remboursement, TauxReduit, Voiture : TCheckBox;
    Annee, Honoraire : string;
    Valid : TToolBarButton97;

    procedure Validation(Sender: TObject);
{$IFDEF COMPTA}
    procedure CalculCumul (Sender : Tobject);
    function ValideRemuneration : boolean;
    procedure MAJAuxi;
{$ENDIF COMPTA}

  end;

  procedure LanceFiche_DADS2Honoraire (LeRang, LeLeQuel, StArguments : String);

implementation

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. : procédure pour appel depuis unités externes
Mots clefs ... :
*****************************************************************}
procedure LanceFiche_DADS2Honoraire (LeRang, LeLeQuel, StArguments : String);
begin
  {$IFDEF COMPTA}
  AGLLanceFiche('PAY', 'DADS2_HONOR_CP', LeRang, LeLeQuel, StArguments);
  {$ELSE}
  AGLLanceFiche('PAY', 'DADS2_HONOR', LeRang, LeLeQuel, StArguments);
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2HONORAIRES.OnNewRecord;
var
IMax : Integer;
QRechNumHonor :TQuery;
StOrdre: string;
begin
Inherited ;
IMax:=1;
// DEB PT8
QRechNumHonor := OpenSQL('SELECT PDH_HONORAIRE FROM DADS2HONORAIRES ORDER BY PDH_HONORAIRE DESC', TRUE);
while not QRechNumHonor.EOF do
      begin
      StOrdre := QRechNumHonor.FindField('PDH_HONORAIRE').AsString;
{PT9
      if ISNumeric(StOrdre) then
         begin
         IMax := StrToInt(StOrdre)+1;
         Break;
         end;
}
      if ((ISNumeric(StOrdre)) and (IMax<=StrToInt(StOrdre))) then
         IMax := StrToInt(StOrdre)+1;
      QRechNumHonor.Next;
      end;
// FIN PT8
Ferme(QRechNumHonor);

StOrdre := FormatFloat('00000000000000000', IMax);
SetField('PDH_HONORAIRE',StOrdre);
SetField('PDH_VALIDITE',Annee);
//SetField('PDH_SECTIONETAB','00');     PT11
SetField('PDH_SECTIONETAB','01');   //  PT11
SetField('PDH_TYPEDADS','0');
SetControlEnabled('PDH_HONORAIRE',  False);
SetControlEnabled('PDH_VALIDITE',  False);
SetField('PDH_AVANTAGENATN', '    ');
SetField('PDH_NTIC', ' ');
SetField('PDH_CHARGEINDEMN', '   ');
SetField('PDH_TAUXSOURCE', '  ');
end;

procedure TOM_DADS2HONORAIRES.OnDeleteRecord;
begin
Inherited ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2HONORAIRES.OnUpdateRecord ;
var
BufChamp, BufDest, BufOrig : string;
StEtab : string;
QRechEtab : TQuery;
begin
Inherited ;
if (GetField('PDH_ETABLISSEMENT')='') then
   SetField('PDH_ETABLISSEMENT', GetParamSoc('SO_ETABLISDEFAUT'));
StEtab:= 'SELECT ET_SIRET'+
         ' FROM ETABLISS WHERE'+
         ' ET_ETABLISSEMENT="'+GetField('PDH_ETABLISSEMENT')+'"';
QRechEtab:= OpenSQL(StEtab,TRUE);
if Not QRechEtab.EOF then
   SetField('PDH_SIRET', QRechEtab.Fields[0].Asstring);
Ferme(QRechEtab);

BufOrig:= GetField ('PDH_SIRET');
{$IFDEF COMPTA}
  EpureChar(BufOrig, BufDest);
  if not VerifSiret(BufDest) then
{$ELSE}
ForceNumerique (BufOrig, BufDest);
if ControlSiret (BufDest)=False then
{$ENDIF}
   begin
   LastError:= 1;
   PgiBox ('Le SIRET de l''établissement#13#10'+
           '"'+GetField ('PDH_ETABLISSEMENT')+'" n''est pas valide',
           Ecran.Caption);
   SetFocusControl ('PDH_ETABLISSEMENT');
   exit;
   end;
{ FQ 20739 BVE 18.06.07 }
  {$IFDEF COMPTA}
  { FQ 20738 BVE 18.06.07 }
  if not(ValideRemuneration) then
  begin         
     LastError:= 1;
     PgiBox('Aucun montant n''est renseigné');
     SetFocusControl('TSREMHONOR');
     Exit;
  end;
  { END FQ 20738 }
  if (GetField ('PDH_RAISONSOCBEN')<>'') then
  begin
     if (GetField ('PDH_SIRETBEN')='') then
        PGIInfo('Le siret du bénéficiaire n''est pas renseigné');
     if (Length(GetField ('PDH_SIRETBEN'))<>14) and (Length(GetField('PDH_SIRETBEN')) > 1) Then
     begin
        LastError:= 1;
        PgiBox ('Le SIRET est incomplet (14 caractères obligatoires)',
                Ecran.Caption);
        SetFocusControl ('PDH_SIRETBEN');
        exit;
     end
     else
     begin
        EpureChar(GetField('PDH_SIRETBEN'), BufDest);
        if (BufDest <> '') and not(VerifSiret(BufDest)) then
        begin
           LastError:= 1;
           PgiBox ('Le SIRET n''est pas valide', Ecran.Caption);
           SetFocusControl ('PDH_SIRETBEN');
           exit;
        end;
     end;
  end;
{$ELSE}
If (GetField ('PDH_RAISONSOCBEN')<>'') then
begin
   if (GetField ('PDH_SIRETBEN')='') then
   begin
      PgiBox ('Le siret du bénéficiaire n''est pas renseigné', Ecran.Caption);
      SetFocusControl ('PDH_SIRETBEN');
      exit;
   end
   else
   begin
      If (Length(GetField('PDH_SIRETBEN'))<>14) Then
      begin
         LastError:= 1;
         PgiBox ('Le SIRET est incomplet (14 caractères obligatoires)',
                 Ecran.Caption);
         SetFocusControl ('PDH_SIRETBEN');
         exit;
      end
      Else
      begin
        If ControlSiret(GetField ('PDH_SIRETBEN'))=False Then
        begin
           LastError:= 1;
           PgiBox ('Le SIRET n''est pas valide', Ecran.Caption);
           SetFocusControl ('PDH_SIRETBEN');
           exit;
        end;
     end;
  end;
end;
{$ENDIF}

{ END FQ 20739 }

If ((GetField ('PDH_RAISONSOCBEN')='') and (GetField ('PDH_NOMBEN')='')) Then
   begin
   LastError:= 1;
   PgiBox ('La raison sociale du bénéficiaire est obligatoire pour une#13#10'+
           'personne morale.#13#10'+
           'Le nom du bénéficiaire est obligatoire pour une personne #13#10'+
           'physique.#13#10', Ecran.Caption);
   SetFocusControl ('PDH_NOMBEN');
   exit;
   end;

If ((GetField ('PDH_RAISONSOCBEN')<>'') and (GetField ('PDH_NOMBEN')<>'')) then
   begin
   PgiBox ('La raison sociale du bénéficiaire et le nom du bénéficiaire#13#10'+
           'ne peuvent pas être renseignés simultanément. La raison #13#10'+
           'sociale étant renseignée, le nom et le prénom sont effacés.',
           Ecran.Caption);
   SetField ('PDH_NOMBEN', '');
   SetField ('PDH_PRENOMBEN', '');
   end;

If ((GetField ('PDH_NOMBEN')<>'') and (GetField ('PDH_PRENOMBEN')='')) Then
   begin
   LastError:= 1;
   PgiBox ('Le nom et le prénom du bénéficiaire sont obligatoires #13#10'+
           'pour une personne physique.#13#10', Ecran.Caption);
   SetFocusControl ('PDH_PRENOMBEN');
   exit;
   end;

If (GetField ('PDH_PROFESSIONBEN')='') Then
   begin
   LastError:= 1;
   PgiBox ('La profession du bénéficiaire n''est pas renseignée',
           Ecran.Caption);
   SetFocusControl('PDH_PROFESSIONBEN');
   exit;
   end;

If (GetField ('PDH_CODEPOSTAL')='') Then
   begin
   LastError:= 1;
   PgiBox ('Le code postal n''est pas renseigné', Ecran.Caption);
   SetFocusControl('PDH_CODEPOSTAL');
   exit;
   end
else
   begin
   If Not (IsNumeric(GetField ('PDH_CODEPOSTAL'))) then
      begin
      LastError:= 1;
      PgiBox ('Le code postal doit être numérique', Ecran.Caption);
      SetFocusControl('PDH_CODEPOSTAL');
      exit;
      end
   Else
      begin
      If StrToInt(GetField ('PDH_CODEPOSTAL'))=0 Then
         begin
         LastError:= 1;
         PgiBox ('Le code postal ne peut pas avoir la valeur 0', Ecran.Caption);
         SetFocusControl('PDH_CODEPOSTAL');
         exit;
         end;
      end;
   end;

If (GetField ('PDH_BUREAUDISTRIB')='') Then
   begin
   LastError:= 1;
   PgiBox ('Le bureau distributeur n''est pas renseigné', Ecran.Caption);
   SetFocusControl('PDH_BUREAUDISTRIB');
   exit;
   end;

if ((Nourriture.Checked=False) and (Logement.Checked=False) and
   (Voiture.Checked=False) and (AuAvant.Checked=False) and
   (NTIC.Checked=False) and (GetField ('PDH_REMAVANTAGE')<>0)) then
   begin
   LastError:= 1;
   PgiBox ('Le montant des avantages en nature est égal à '+
           FloatToStr(GetField('PDH_REMAVANTAGE'))+'#13#10'+
           'mais aucune zone correspondante n''est cochée', Ecran.Caption);
   SetFocusControl ('PDH_REMAVANTAGE');
   exit;
   end;

BufChamp:= GetField('PDH_AVANTAGENATN');
if ((Nourriture<>nil) and (Logement<>nil) and (Voiture<>nil) and
   (AuAvant<>nil)) then
   if (((Copy (BufChamp,1,1)='N')<>Nourriture.Checked) or
      ((Copy (BufChamp,2,1)='L')<>Logement.Checked) or
      ((Copy (BufChamp,3,1)='V')<>Voiture.Checked) or
      ((Copy (BufChamp,4,1)='A')<>AuAvant.Checked)) then
      begin
      if (Nourriture.Checked=TRUE) then
         BufChamp:= 'N'
      else
         BufChamp:= ' ';

      if (Logement.Checked=TRUE) then
         BufChamp:= BufChamp+'L'
      else
         BufChamp:= BufChamp+' ';

      if (Voiture.Checked=TRUE) then
         BufChamp:= BufChamp+'V'
      else
         BufChamp:= BufChamp+' ';

      if (AuAvant.Checked=TRUE) then
         BufChamp:= BufChamp+'A'
      else
         BufChamp:= BufChamp+' ';
      SetField ('PDH_AVANTAGENATN', BufChamp);
      end;

BufChamp:= GetField ('PDH_NTIC');
if (NTIC<>nil) then
   if ((BufChamp='T')<>NTIC.Checked) then
      begin
      if (NTIC.Checked=TRUE) then
         BufChamp:= 'T'
      else
         BufChamp:= ' ';
      SetField ('PDH_NTIC', BufChamp);
      end;

if ((AlloForfait.Checked=False) and (Remboursement.Checked=False) and
   (Employeur.Checked=False) and (GetField ('PDH_REMINDEMNITE')<>0)) then
   begin
   LastError:= 1;
   PgiBox ('Le montant des frais professionnels est égal à '+
           FloatToStr(GetField('PDH_REMINDEMNITE'))+'#13#10'+
           'mais aucune zone correspondante n''est cochée', Ecran.Caption);
   SetFocusControl ('PDH_REMINDEMNITE');
   exit;
   end;

BufChamp:= GetField('PDH_CHARGEINDEMN');
if ((AlloForfait<>nil) and (Remboursement<>nil) and (Employeur<>nil)) then
   if (((Copy (BufChamp, 1, 1)='F')<>AlloForfait.Checked) or
      ((Copy (BufChamp, 2, 1)='R')<>Remboursement.Checked) or
      ((Copy (BufChamp, 3, 1)='P')<>Employeur.Checked)) then
      begin
      if (AlloForfait.Checked=TRUE) then
         BufChamp:= 'F'
      else
         BufChamp:= ' ';

      if (Remboursement.Checked=TRUE) then
         BufChamp:= BufChamp+'R'
      else
         BufChamp:= BufChamp+' ';

      if (Employeur.Checked=TRUE) then
         BufChamp:= BufChamp+'P'
      else
         BufChamp:= BufChamp+' ';
      SetField('PDH_CHARGEINDEMN', BufChamp);
      end;

BufChamp:= GetField('PDH_TAUXSOURCE');
if ((TauxReduit<>nil) and (Dispense<>nil)) then
   if (((Copy (BufChamp, 1, 1)='R')<>TauxReduit.Checked) or
      ((Copy (BufChamp, 2, 1)='D')<>Dispense.Checked)) then
      begin
      if (TauxReduit.Checked=TRUE) then
         BufChamp:= 'R'
      else
         BufChamp:= ' ';

      if (Dispense.Checked=TRUE) then
         BufChamp:= BufChamp+'D'
      else
         BufChamp:= BufChamp+' ';
      SetField('PDH_TAUXSOURCE', BufChamp);
      end;
{ FQ 20733 BVE 05.09.07 }
{$IFDEF COMPTA}
  { On doit mettre à jour les coordonnées de la Table TIERS
    PDH_SIRETBEN <==> T_SIRET
    PDH_NOMBEN <==> T_LIBELLE
    PDH_PRENOMBEN <==> T_PRENOM
    PDH_RAISONSOCBEN <==> T_LIBELLE
    PDH_ADRCOMPL <==> T_ADRESSE2 + T_ADRESSE3
    PDH_ADRNUM + PDH_ADRNOM <==> T_ADRESSE1
    PDH_CODEPOSTAL <==> T_CODEPOSTAL
    PDH_BUREAUDISTRIB <==> T_VILLE }
  
{$ENDIF}
{ END FQ 20733 }
end ;

procedure TOM_DADS2HONORAIRES.OnAfterUpdateRecord;
begin
Inherited ;
{$IFDEF COMPTA}
{ FQ 20728 BVE 12.09.07 }
  if ExisteSQL('SELECT 1 FROM TIERS WHERE T_AUXILIAIRE = "' + GetControlText('PDH_HONORAIRE') + '"') then
  begin
     MAJAuxi;
  end;
{END FQ 20728 }
{$ENDIF}
end;


{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/09/2007
Modifié le ... : 12/09/2007
Description .. : Permet de mettre à jour l'enregistrement contenu dans la
Suite ........ : table DADSHONORAIRES à partir des données qui
Suite ........ : viennent d'être saisies
Suite ........ : FQ 20733
Mots clefs ... :
*****************************************************************}
procedure TOM_DADS2HONORAIRES.MAJAuxi;
var
  TobAuxi, TobDads2 : TOB;
  Q  : TQuery;
  SQL: String;
begin
  try
     TobDads2 := TOB.Create('DADS2HONORAIRES',nil,-1);
     TobAuxi  := TOB.Create('TIERS',nil,-1);

     // Chargement des données
     TobDads2.GetEcran(ecran);
     SQL := 'SELECT * FROM TIERS WHERE T_AUXILIAIRE = "' + GetControlText('PDH_HONORAIRE') + '"';
     Q := OpenSQL(SQL,false);
     while not Q.Eof do
     begin
        TobAuxi.SelectDB('',Q);
        Dads2toTiers(TobDads2,TobAuxi);
        TobAuxi.UpdateDB();
        Q.Next;
     end;
  finally
     if assigned(TobAuxi) then FreeAndNil(TobAuxi);
     if assigned(TobDads2) then FreeAndNil(TobDads2);
     if assigned(Q) then Ferme(Q);
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2HONORAIRES.OnLoadRecord;
var
Buf, BufAn, BufJour, BufMois : string;
AnneeA, Jour, MoisM : Word;
begin
Inherited ;
DecodeDate(DebExer, AnneeA, MoisM, Jour);
BufJour:='';
BufMois:='';
BufAn:='';
Buf := GetField('PDH_AVANTAGENATN');
{PT3
if ((Copy(Buf, 1, 1) = 'N') and (Nourriture <> nil)) then
   Nourriture.Checked := TRUE;
if ((Copy(Buf, 2, 1) = 'L') and (Logement <> nil)) then
   Logement.Checked := TRUE;
if ((Copy(Buf, 3, 1) = 'V') and (Voiture <> nil)) then
   Voiture.Checked := TRUE;
if ((Copy(Buf, 4, 1) = 'A') and (AuAvant <> nil)) then
   AuAvant.Checked := TRUE;
}
if (Nourriture <> nil) then
   begin
   if (Copy(Buf, 1, 1)='N') then
      Nourriture.Checked:= TRUE
   else
      Nourriture.Checked:= FALSE;
   end;
if (Logement <> nil) then
   begin
   if (Copy(Buf, 2, 1)='L') then
      Logement.Checked:= TRUE
   else
      Logement.Checked:= FALSE;
   end;
if (Voiture <> nil) then
   begin
   if (Copy(Buf, 3, 1)='V') then
      Voiture.Checked:= TRUE
   else
      Voiture.Checked:= FALSE;
   end;
if (AuAvant <> nil) then
   begin
   if (Copy(Buf, 4, 1)='A') then
      AuAvant.Checked:= TRUE
   else
      AuAvant.Checked:= FALSE;
   end;
//FIN PT3

//PT1-1
Buf := GetField('PDH_NTIC');
{PT3
if ((Buf='T') and (NTIC <> nil)) then
   NTIC.Checked := TRUE;
}
if (NTIC <> nil) then
   begin
   if (Buf='T') then
      NTIC.Checked:= TRUE
   else
      NTIC.Checked:= FALSE;
   end;
//FIN PT3
//FIN PT1-1

Buf := GetField('PDH_CHARGEINDEMN');
{PT3
if ((Copy(Buf, 1, 1) = 'F') and (AlloForfait <> nil)) then
   AlloForfait.Checked := TRUE;
if ((Copy(Buf, 2, 1) = 'R') and (Remboursement <> nil)) then
   Remboursement.Checked := TRUE;
if ((Copy(Buf, 3, 1) = 'P') and (Employeur <> nil)) then
   Employeur.Checked := TRUE;
}
if (AlloForfait <> nil) then
   begin
   if (Copy(Buf, 1, 1)='F') then
      AlloForfait.Checked:= TRUE
   else
      AlloForfait.Checked:= FALSE;
   end;
if (Remboursement <> nil) then
   begin
   if (Copy(Buf, 2, 1)='R') then
      Remboursement.Checked:= TRUE
   else
      Remboursement.Checked:= FALSE;
   end;
if (Employeur <> nil) then
   begin
   if (Copy(Buf, 3, 1)='P') then
      Employeur.Checked:= TRUE
   else
      Employeur.Checked:= FALSE;
   end;
//FIN PT3

Buf := GetField('PDH_TAUXSOURCE');
{PT3
if ((Copy(Buf, 1, 1) = 'R') and (TauxReduit <> nil)) then
   TauxReduit.Checked := TRUE;
if ((Copy(Buf, 2, 1) = 'D') and (Dispense <> nil)) then
   Dispense.Checked := TRUE;
}
if (TauxReduit <> nil) then
   begin
   if (Copy(Buf, 1, 1)='R') then
      TauxReduit.Checked:= TRUE
   else
      TauxReduit.Checked:= FALSE;
   end;
if (Dispense <> nil) then
   begin
   if (Copy(Buf, 2, 1)='D') then
      Dispense.Checked:= TRUE
   else
      Dispense.Checked:= FALSE;
   end;
//FIN PT3
end;

procedure TOM_DADS2HONORAIRES.OnChangeField(F: TField);
begin
Inherited ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2HONORAIRES.OnArgument(S: string);
var
Pages : TPageControl;
Buf : string;
{$IFDEF COMPTA}
  Nbc                   : Integer;
  i                     : Integer;
  LeChamp               : String;
  Ledit                 : THEdit;
{$ENDIF}

begin
Inherited ;
Buf:=Trim(ReadTokenPipe(S,';')) ;
Honoraire:=Trim(ReadTokenPipe(S,';')) ;
Annee:=Trim(ReadTokenPipe(S,';')) ;

Nourriture := TCheckBox (GetControl ('CHNOURRITURE'));
{PT4
if (Nourriture <> nil) then
   Nourriture.OnExit:=CExit;
}
Logement := TCheckBox (GetControl ('CHLOGEMENT'));
{PT4
if (Logement <> nil) then
   Logement.OnExit:=CExit;
}
Voiture := TCheckBox (GetControl ('CHVOITURE'));
{PT4
if (Voiture <> nil) then
   Voiture.OnExit:=CExit;
}
AuAvant := TCheckBox (GetControl ('CHAUTRESAVANTAGES'));
{PT4
if (AuAvant <> nil) then
   AuAvant.OnExit:=CExit;
}
//PT1-1
NTIC := TCheckBox (GetControl ('CHNTIC'));
{PT4
if (NTIC <> nil) then
   NTIC.OnExit:=CExit;
}
//FIN PT1-1

AlloForfait := TCheckBox (GetControl ('CHALLOCFORFAIT'));
{PT4
if (AlloForfait <> nil) then
   AlloForfait.OnExit:=CExit;
}
Remboursement := TCheckBox (GetControl ('CHREMBOURSEMENTS'));
{PT4
if (Remboursement <> nil) then
   Remboursement.OnExit:=CExit;
}
Employeur := TCheckBox (GetControl ('CHEMPLOYEUR'));
{PT4
if (Employeur <> nil) then
   Employeur.OnExit:=CExit;
}

TauxReduit := TCheckBox (GetControl ('CHTAUXREDUIT'));
{PT4
if (TauxReduit <> nil) then
   TauxReduit.OnExit:=CExit;
}
Dispense := TCheckBox (GetControl ('CHDISPENSE'));
{PT4
if (Dispense <> nil) then
   Dispense.OnExit:=CExit;
}

// Positionnement sur le premier onglet
Pages := TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

//PT4
// Gestion du navigateur
Valid := TToolBarButton97(GetControl('BValider'));
if Valid<>nil then
   Valid.OnClick:=Validation;
//FIN PT4


{$IFDEF COMPTA}
  Nbc                   := Ecran.ComponentCount -1;
  for i := 0 to Nbc do
  begin
    // on ne prend que les Thedit pour saisie
    if (Ecran.Components[i].ClassName = 'THEdit') and (Ecran.Components[i].Tag = 1) then
    begin
      LeChamp           := Ecran.Components[i].Name;
      Ledit             := THEdit(GetControl(LeChamp));
      Ledit.OnExit      := CalculCumul;
      Ledit.Text        := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
    end;

    if (Ecran.Components[i].ClassName = 'THDBEdit') and (Ecran.Components[i].Tag = 2) then
    begin
      LeChamp           := Ecran.Components[i].Name;
      Ledit             := THEdit(Getcontrol(LeChamp));
      Ledit.Text        := STRFMONTANT(Valeur(Ledit.Text), 15, V_PGI.OkDecV, '', True);
      Ledit.Enabled     := False;
    end;
  end;
{$ENDIF}

end;

procedure TOM_DADS2HONORAIRES.OnClose;
begin
Inherited ;
end;

procedure TOM_DADS2HONORAIRES.OnCancelRecord;
begin
Inherited ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/08/2003
Modifié le ... :   /  /
Description .. : Click sur une donnée autre que DB
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
{PT4
procedure TOM_DADS2HONORAIRES.CExit(Sender: TObject);
begin
if not (ds.state in [dsinsert,dsedit]) then ds.edit;
end;
}


//PT4
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/12/2004
Modifié le ... :   /  /
Description .. : Validation
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2HONORAIRES.Validation(Sender: TObject);
begin
if not (ds.state in [dsinsert,dsedit]) then
   ds.edit;
TFFiche(Ecran).BValiderClick(Nil);
end;
//FIN PT4

{$IFDEF COMPTA}
procedure TOM_DADS2HONORAIRES.CalculCumul(Sender : Tobject);
var
  LeChamp               : String;
  Cumul                 : Extended;
  LaValeur              : Extended;


begin
  Lechamp               := THEdit(sender).Name;
  LaValeur              := Valeur(THEdit(Sender).Text);
  LeChamp               := 'PDH_' + copy(LeChamp, 3, length(LeChamp));
  Cumul                 := Valeur(GetField(LeChamp));
  Cumul                 := Cumul + LaValeur;
  SetField(LeChamp, STRFMONTANT(Cumul, 15, V_PGI.OkDecV, '', True));
  THEdit(Sender).Text := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 18/06/2007
Modifié le ... : 18/06/2007
Description .. : On vérifie que des données on bien été enregistrées dans 
Suite ........ : l'onglet rémunération.
Suite ........ : FQ 20738
Mots clefs ... : 
*****************************************************************}
function TOM_DADS2HONORAIRES.ValideRemuneration : boolean;
var
  ListeComposants : string;
  Composant       : string;
begin
  result := false;
  ListeComposants := 'PDH_REMHONOR;PDH_REMCOMMISS;PDH_REMCOURTAGE;PDH_REMRISTOURNE;' +
                     'PDH_REMJETON;PDH_REMAUTEUR;PDH_TVAAUTEUR;PDH_REMINVENT;'+
                     'PDH_REMAUTRE;PDH_RETENUESOURC;PDH_REMINDEMNITE;PDH_REMAVANTAGE';
  Composant := ReadTokenST(ListeComposants);
  while (Composant <> '') do
  begin
     if Valeur(GetField(Composant)) <> 0 then
     begin
        Result := true;
        Break;
     end;     
     Composant := ReadTokenST(ListeComposants);
  end;
end;
{$ENDIF COMPTA}

initialization
  registerclasses([TOM_DADS2HONORAIRES]);
end.

