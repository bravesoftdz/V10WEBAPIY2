{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /
Description .. : Fusion Word
Mots clefs ... : PAIE;WORD
*****************************************************************
PT1   : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
      : 07/09/2001 SB V547 Fiche de bug 283
PT2-1 :                    Nouvelle fonction de recherche d'un libélle associé à
                           un code
PT2-2 :                    Chargement des champs ParamSoc
PT3   : 19/10/2001 SB V562 Fiche de bug 343
                           Modification code retour Devise
PT4   : 13/12/2001 SB V563 Fn GetLibelle ne renvoie pas forcement la bonne
                           tablette
PT5   : 10/01/2002 SB V571 Contrôle de cohérance
PT6   : 29/03/2002 JL V571 Ajout médecine du travail
PT7   : 08/11/2002 JL V595 Ajout Formation
PT8   : 12/12/2002 SB V585 Amélioration traitement function GetLibelle
PT9   : 17/02/2003 PH V_42 Suppression des fonctions vides et non utilisées
PT10  : 20/02/2003 SB V_42 Formatage du n°ss érroné par FormateNombre si
                           alphanumérique
PT11  : 05/08/2003 SB V_42 FQ 10681 Formatage du nom du fichier de sortie
PT12  : 24/02/2003 SB V_50 FQ 11130 CalculDuree érronné si salarié entrée sortie
                           le même jour
PT13  : 23/04/2004 SB V_50 FQ 11084 Nouvelle fonction : DUREECONTRATJOUR
PT14  : 28/04/2004 SB V_50 FQ 10812 Intégration de la gestion des déclarants
PT15  : 28/04/2004 SB V_50 FQ 11085 Ajout PSA_E pour accord sur les textes
                           reliées au salarié
PT16  : 14/05/2004 SB V_50 FQ 11301 Le getvar à Null plante la conversion
PT17  : 07/02/2005 SB V_60 FQ 11785 Ajout fn renvoie durée préavis @DUREEPREAVIS
PT18  : 21/03/2005 SB V_60 FQ 11423 Ajout fn renvoie donnée organisme
PT19  : 21/03/2005 SB V_60 Transtypation code en fonction Public
PT20  : 26/04/2005 SB V_60 FQ 12067 Plus un mois calcul fonction duree contrat
PT21  : 16/09/2005 SB V_65 FQ 12456 Ajout fn élement de la paie
PT22  : 20/09/2005 SB V_65 FQ 12421 Ajout info registre, table annuaire
PT23  : 08/11/2005 SB V_65 FQ 12675 Fn element de paie renvoie 0 si pas de cumul
                           ou rub. renseigné
PT24  : 02/06/2006 SB V_65 Optimisation mémoire
PT25  : 11/01/2007 GGU V_80 FQ 13478 durée de travail erronée
PT26  : 17/04/2007 FC V_72 FQ 13954 Rajout d'une fonction DUREECONTRATDEBCTR qui
                           se base sur la date de début contrat
PT27  : 10/05/2007 FC V_72 Rajout de la fonction GETELTNAT qui récupère la
                           valeur d'un élément national
PT28  : 27/06/2007 FC V_72 FQ 13953 Rajout des fonctions DUREEESSAISEMAINE,
                           DUREEPREAVISSEMAINE, DUREECONTRATSEMAINE
PT29  : 05/09/2007 FC V_80 FQ14582 Fonction @getcalendrier : ne pas éditer 0h 0h
PT31 08/07/2008 NA V_850 FQ 11785 Ajout fn renvoie durée du préavis en jours @DUREEPREAVISJOUR
}
unit PGFusionWord;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  SysUtils, Classes, Controls, Forms,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Doc_Parser, AGLUtilOLE, Utob, HCtrls,UyFileSTD,
  HMsgBox, HEnt1, PgOutils2,
  ulibEditionPaie, // pt31
  ParamSoc;  //PT27
type
  TFFusion = class(TForm)
    procedure OnInitialization(Sender: TObject);
    procedure OnFinalization(Sender: TObject);
    procedure OnGetList(Sender: TObject; IdentList: string; var List: TStringList);
    procedure OnGetVar(Sender: TObject; VarName: string; VarIndx: Integer; var Value: Variant);
    procedure OnSetVar(Sender: TObject; VarName: string; VarIndx: Integer; Value: variant);
    procedure OnFunction(Sender: TObject; FuncName: string; Params: array of variant; var Result: Variant);
  private { Déclarations privées } Tob_Salaries, Tob_Net, Tob_Calendrier, Tob_Contrat, Tob_Societe: tob;
    Tob_Convention, Tob_Organisme, Tob_Medecine, Tob_VisiteMedicale, Tob_ConvocFormation: Tob;
    Objet, path: string;
    Argument: TStringList;
    DateDebPaie, DateFinPaie: TDateTime;
    Tob_Declarant,Tob_Annuaire : Tob; { PT14 PT22 }
    procedure LanceFusion;
    function GetCodeSalarie(iIndx: integer): string;
    //  function GetESexe(Sexe: string): string;
    function GetCalendrier(VSalarie: string): string;
    function RecupValCodeEltDos(CodeElt,CodeSal,TypeNiveau,ValeurNiveau,StDD : String) : String;
    function RecupValCodeEltNat(CodeElt,CodeSal,Predefini,StDD : String) : String;
    function DetermineAlsaceOuNon(CodeSal : String; Montant, MontantEuro : Double): String;
    //  Function GetTablettelibelle(Champ,value : String) : String;
  public { Déclarations publiques }
  end;

procedure LanceFusionWord(Obj, Pat: string; Salarie: TStringList; T_Sal, T_Net, T_Cal, T_Contrat, T_Soc, T_Conv, T_Org, T_VisiteMed, T_AdrMed, T_ConvocForm,T_Annuaire: tob);
procedure ValidateScript(theScript: string; var bValid: boolean);
var FileP : String;

implementation


uses EntPaie, DB;

{$R *.DFM}

procedure LanceFusionWord(Obj, Pat: string; Salarie: TStringList; T_Sal, T_Net, T_Cal, T_Contrat, T_Soc, T_Conv, T_Org, T_VisiteMed, T_AdrMed, T_ConvocForm,T_Annuaire: tob); //PT6 et PT7
var
  FFusion: TFFusion;
begin
  FFusion := TFFusion.Create(Application);
  try
    FFusion.Argument := Salarie;
    FFusion.Objet := Obj;
    FFusion.path := Pat;
    FFusion.Tob_Salaries := T_Sal;
    FFusion.Tob_Net := T_Net;
    FFusion.Tob_Calendrier := T_Cal;
    FFusion.Tob_Contrat := T_Contrat;
    FFusion.Tob_Societe := T_Soc;
    FFusion.Tob_Convention := T_Conv;
    FFusion.Tob_Organisme := T_Org;
    FFusion.Tob_VisiteMedicale := T_VisiteMed;
    FFusion.Tob_Medecine := T_AdrMed;
    FFusion.Tob_ConvocFormation := T_ConvocForm;
    FFusion.Tob_Annuaire := T_Annuaire;         { PT22 }
    FFusion.LanceFusion;
  finally
    if FFusion <> nil then FFusion.Free;
  end;
end;

procedure TFFusion.LanceFusion;
var
  NomFichier, NomFichierSortie: string;
  bOk: boolean;
  iRet: integer;
  CodeRetour : Integer;
begin
  if (VH_Paie.PGCheminRech <> '') and (VH_Paie.PGCheminSav <> '') then
  begin
    if Objet = 'CERTIFICAT' then
    begin
      NomFichier := VH_Paie.PGCheminRech + '\CertificatTravail.doc';
      NomFichierSortie := VH_Paie.PGCheminSav + '\' + FormatDateTime('dmyy_hns_z', Now) + '_Certificat.doc'; //PT11
    end
    else
      if Objet = 'SOLDE' then
    begin
      NomFichier := VH_Paie.PGCheminRech + '\Soldedetoutcompte.doc';
      NomFichierSortie := VH_Paie.PGCheminSav + '\' + FormatDateTime('dmyy_hns_z', Now) + '_Solde.doc'; //PT11
    end
    else
      if Objet = 'CONTRAT' then
    begin
      NomFichier := Path;
      NomFichierSortie := VH_Paie.PGCheminSav + '\' + FormatDateTime('dmyy_hns_z', Now) + '_Contrat.doc'; //PT11
    end
    else
      if Objet = 'MEDECINE' then //PT6
    begin
      NomFichier := VH_Paie.PGCheminRech + '\MedecineTravail.doc';
      NomFichierSortie := VH_Paie.PGCheminSav + '\' + FormatDateTime('dmyy_hns_z', Now) + '_Medecine.doc'; //PT11
    end
    else
      if Objet = 'FORMATION' then //PT7
    begin
      NomFichier := VH_Paie.PGCheminRech + '\ConvocForm.doc';
      NomFichierSortie := VH_Paie.PGCheminSav + '\' + FormatDateTime('dmyy_hns_z', Now) + '_Formation.doc'; //PT11
    end;
  end;
  If FileP <> '' then NomFichier := FileP;
  bOk := true;
  iret := -1;
  try
    iRet := ConvertDocFile(NomFichier, NomFichierSortie, OnInitialization, OnFinalization, OnFunction, OnGetVar, OnSetVar, OnGetList);
  except
    bOk := false;
  end;
  if BOk = False then
    PGIBox('La conversion du document n''a pu aboutir.', 'Abandon du traitement');
  if (FileExists(NomFichierSortie)) = False then
    PgiBox('Le fichier de sortie "' + NomFichierSortie + '" n''existe pas.', 'Abandon du traitement');

  if (iRet = 0) and (bOk) and (FileExists(NomFichierSortie)) then
    OpenDoc(NomFichierSortie, wdWindowStateMaximize, wdPrintView);
   //suppression des fichiers exportés sur le disque local
   If FileP <> '' then
   begin
        If FileExists(Pchar(FileP)) then
              DeleteFile(pchar(FileP));
   end;
end;

procedure TFFusion.OnGetVar(Sender: TObject; VarName: string; VarIndx: Integer; var Value: Variant);
var
  VSalarie, VEtab, VMedecine, St: string;
  TMed, TRech, TSal: Tob;
  Net: Double;
begin
  Value := '';
  VSalarie := '';
  { DEB PT14 }
  if Pos('PAY_DECLARANT', VarName) > 0 then
  begin
    St := Copy(VarName, 15, Length(VarName));
    if Assigned(Tob_Declarant) then
    begin
      TRech := Tob_Declarant.FindFirst(['ORDRE'], [VarIndx], False);
      if Assigned(TRech) then
        if TRech.FieldExists(St) then Value := TRech.GetValue(St);
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  { FIN PT14 }

  if VarName = 'NBSAL' then
  begin
    value := Argument.count;
    exit;
  end;

  if VarIndx <> -1 then VSalarie := (GetCodeSalarie(VarIndx)); // else VSalarie='') then exit;
  if VSalarie <> '' then TSal := Tob_Salaries.FindFirst(['PSA_SALARIE'], [VSalarie], False)
  else TSal := nil;
  if (TSal <> nil) and (VSalarie <> '') then
  begin
    VEtab := TSal.GetValue('PSA_ETABLISSEMENT');
    if (Pos('PSA_', VarName) = 1) or (Pos('ET_', VarName) = 1) or (Pos('ETB_', VarName) = 1) then
    begin
      value := TSal.GetValue(VarName); //GetTabletteLibelle(Varname,);
      { DEB PT10 Formatage du n°SS }
      if VarName = 'PSA_NUMEROSS' then
        Value := Copy(Value, 1, 1) + ' ' + Copy(Value, 2, 2) + ' ' + Copy(Value, 4, 2) + ' ' +
          Copy(Value, 6, 2) + ' ' + Copy(Value, 8, 3) + ' ' + Copy(Value, 11, 3) + ' ' +
          Copy(Value, 14, 2);
      { FIN PT10 }
      { DEB PT15 }
      if VarName = 'PSA_E' then
        if TSal.GetValue('PSA_SEXE') = 'F' then Value := 'e';
      { FIN PT15 }
      if Value = null then value := ''; { PT16 }
      exit;
    end;
  end;
  //RecupValeur Dernier Contrat Tob contrat de travail
  if (Pos('PCI_', VarName) = 1) and (Tob_Contrat <> nil) and (VSalarie <> '') then
  begin
    TRech := Tob_Contrat.FindFirst(['PCI_SALARIE'], [VSalarie], False);
    if Trech <> nil then
    begin
      value := Trech.GetValue(VarName); //GetTabletteLibelle(Varname,);
      //TRech.free; Trech:=nil;
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;

  //RecupValeur médecine du tarvail       //DEBUT PT6
  if (Pos('PVM_', VarName) = 1) and (Tob_VisiteMedicale <> nil) and (VSalarie <> '') then
  begin
    TRech := Tob_VisiteMedicale.FindFirst(['PVM_SALARIE'], [VSalarie], False);
    if Trech <> nil then
    begin
      //If VarName='PVM_TYPEVISITMED' Then Value:=RechDom('PGTYPVISITEMED',Trech.GetValue(VarName),False)
      //Else value:=Trech.GetValue(VarName);
      value := Trech.GetValue(VarName);
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //RecupValeur adresse médecine du tarvail
  if (Pos('ANN_', VarName) = 1) then
  begin        { DEB PT22 }
    if (VarName = 'ANN_SIREN') OR (VarName = 'ANN_RM') OR (VarName = 'ANN_RMDEP') then
      Begin
      if Assigned(Tob_Annuaire) then TRech := Tob_Annuaire.FindFirst(['DOS_NODOSSIER'], [V_Pgi.NoDossier], False)
      else                           FreeAndNil(TRech);
      if Assigned(TRech) then value := Trech.GetValue(VarName);
      if Value = null then value := ''; 
      exit;
      End      { FIN PT22 }
    else
    if (Tob_VisiteMedicale <> nil) AND (VSalarie <> '') then
      Begin
      TMed := Tob_VisiteMedicale.FindFirst(['PVM_SALARIE'], [VSalarie], False);
      VMedecine := TMed.GetValue('PVM_MEDTRAVGU');
      if (Tob_Medecine <> nil) and (VMedecine <> '') then
        begin
        TRech := Tob_Medecine.FindFirst(['ANN_GUIDPER'], [VMedecine], False);
        if Trech <> nil then
          begin
          value := Trech.GetValue(VarName);
          end;
        if Value = null then value := ''; { PT16 }
        exit;
        end;
      End;
  end; //FIN PT6
  if (Pos('PFO_', VarName) = 1) and (Tob_ConvocFormation <> nil) and (VSalarie <> '') then //PT7
  begin
    TRech := Tob_ConvocFormation.FindFirst(['PFO_SALARIE'], [VSalarie], False);
    if Trech <> nil then
    begin
      value := Trech.GetValue(VarName);
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //RecupValeur Tob convention collective
  if (Pos('PCV_', VarName) = 1) and (Tob_Convention <> nil) and (VSalarie <> '') then
  begin
    TRech := Tob_Convention.FindFirst(['PSA_SALARIE'], [VSalarie], False);
    if Trech <> nil then
    begin
      value := Trech.GetValue(VarName); //GetTabletteLibelle(Varname,);
      //TRech.free; Trech:=nil;
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //RecupValeur Tob Organismes et caisses
  if (Pos('POG_', VarName) = 1) and (Tob_Organisme <> nil) and (VEtab <> '') then
  begin
    TRech := Tob_Organisme.FindFirst(['POG_ORGANISME', 'POG_ETABLISSEMENT'], ['001', VEtab], False); { PT18 }
    if Trech <> nil then
    begin
      value := Trech.GetValue(VarName); //GetTabletteLibelle(Varname,Trech.GetValue(VarName));
      //TRech.free; Trech:=nil;
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //RecupValeur Tob Société
  if (Pos('SO_', VarName) = 1) and (Tob_Societe.detail.count > 0) then
  begin
    TRech := Tob_Societe.FindFirst(['SOC_NOM'], [VarName], False); //PT2-2
    if Trech <> nil then
      value := Trech.GetValue('SOC_DATA'); //GetTabletteLibelle(Varname,);
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //Mode de reglement
  if (VarName = 'PPU_MODE') and (VSalarie <> '') then
  begin
    value := '';
    DateDebPaie := IDate1900;
    DateFinPaie := Idate1900;
    TRech := Tob_Net.FindFirst(['PPU_SALARIE'], [VSalarie], False);
    value := 0;
    while Trech <> nil do
    begin
      if (TRech.GetValue('DATEDEBUT') > DateDebPaie) or (TRech.GetValue('DATEFIN') > DateFinPaie) then
      begin
        DateDebPaie := TRech.GetValue('DATEDEBUT');
        DateFinPaie := TRech.GetValue('DATEFIN');
        value := TRech.GetValue('PPU_PGMODEREGLE'); {PT1}
      end;
      TRech := Tob_Net.FindNext(['PPU_SALARIE'], [VSalarie], False);
    end;
    value := AnsiLowerCase(RechDom('PGMODEREGLE', value, FALSE));
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //Dernier Net à payer
  if (VarName = 'PHC_MONTANT') and (Tob_Net <> nil) and (VSalarie <> '') then
  begin
    DateDebPaie := IDate1900;
    DateFinPaie := Idate1900;
    TRech := Tob_Net.FindFirst(['PPU_SALARIE'], [VSalarie], False);
    value := 0;
    while Trech <> nil do
    begin
      if (TRech.GetValue('DATEDEBUT') > DateDebPaie) or (TRech.GetValue('DATEFIN') > DateFinPaie) then
      begin
        DateDebPaie := TRech.GetValue('DATEDEBUT');
        DateFinPaie := TRech.GetValue('DATEFIN');
        Net := TRech.GetValue('PPU_CNETAPAYER');
        value := Net;
      end;
      TRech := Tob_Net.FindNext(['PPU_SALARIE'], [VSalarie], False);
    end;
    if Value = null then value := ''; { PT16 }
    exit;
  end;
  //Devise
  if VarName = 'DEVISE' then
  begin
    if VH_Paie.PGTenueEuro = True then Value := 'Euro'
    else Value := RechDom('TTDEVISETOUTES', VH_Paie.PGMonnaieTenue, False);
    if value = 'francs' then Value := 'franc' //PT3 Ajout Ligne
    else if value = '' then value := 'Monnaie non communiquée';
    if Value = null then value := ''; { PT16 }
    exit;
  end;

  if TSal <> nil then TSal.free;

end;


procedure TFFusion.OnFinalization(Sender: TObject);
begin
  if Tob_Salaries <> nil then
  begin
    Tob_Salaries.free;
    Tob_Salaries := nil
  end;
  if Tob_Net <> nil then
  begin
    Tob_Net.free;
    Tob_Net := nil
  end;
  if Tob_Calendrier <> nil then
  begin
    Tob_Calendrier.free;
    Tob_Calendrier := nil
  end;
  if Tob_Contrat <> nil then
  begin
    Tob_Contrat.free;
    Tob_Contrat := nil
  end;
  if Tob_Societe <> nil then
  begin
    Tob_Societe.free;
    Tob_Societe := nil
  end;
  if Tob_Convention <> nil then
  begin
    Tob_Convention.free;
    Tob_Convention := nil
  end;
  if Tob_Organisme <> nil then
  begin
    Tob_Organisme.free;
    Tob_Organisme := nil
  end;
  if Tob_VisiteMedicale <> nil then
  begin
    Tob_VisiteMedicale.free;
    Tob_VisiteMedicale := nil
  end; //PT6
  if Tob_medecine <> nil then
  begin
    Tob_Medecine.free;
    Tob_Medecine := nil
  end;
  if Tob_ConvocFormation <> nil then
  begin
    Tob_ConvocFormation.free;
    Tob_ConvocFormation := nil
  end; //PT7
  if Assigned(Tob_Declarant) then FreeAndNil(Tob_Declarant); { PT14 }

  if Assigned(Tob_Annuaire) then FreeAndNil(Tob_Annuaire);  { PT24 }
end;
{
function TFFusion.GetESexe(Sexe : string) : string;
begin
if Sexe='M' then result := '' else result:='e';
end;
}

procedure TFFusion.OnFunction(Sender: TObject; FuncName: string;
  Params: array of variant; var Result: Variant);
var
  ordre: integer;
  T, TRech: TOB;
  CodeSal, NomChamp, ValChamp, Tablette, Etab, Org, StDD, StDF, Cumul : string;
  CodeElt, TypeNiveau, ValeurNiveau : string; //PT27
  SQL, StNat, StChamp, StCum, StCumul : String;
  Q: TQuery;
  DateDeb, Date1, Date2: TDateTime;
  PremMois, PremAnnee, duree: Word;
  nbmois, nbjour : integer; // pt31
begin
  Result := 0;
  T := nil;
  if Length(Params) < 1 then exit; //PT5
if ((FuncName='@GETCALENDRIER') or (FuncName='@DUREECONTRAT') or
   (FuncName='@DUREECONTRATJOUR') or (FuncName='@DUREEESSAI') or
   (FuncName='@DUREEPREAVIS') or (FuncName='@DUREECONTRATDEBCTR') or
   (FuncName='@DUREECONTRATSEMAINE') or (FuncName='@DUREEESSAISEMAINE') or
// pt31 (FuncName = '@DUREEPREAVISSEMAINE')) then //PT26
   (FuncName='@DUREEPREAVISSEMAINE') or (FuncName='@DUREEPREAVISJOUR')) then  // pt31
    { PT13 PT17 }
  begin
    CodeSal := (Params[1]); //GetCodeSalarie
    if (CodeSal = '') then exit;
    T := Tob_Salaries.FindFirst(['PSA_SALARIE'], [CodeSal], False);
    if not assigned(T) then exit;
  end;

  if FuncName = '@GETCALENDRIER' then
  begin
    result := GetCalendrier(CodeSal);
    exit;
  end;

  if FuncName = '@DUREECONTRAT' then
  begin
    if T = nil then exit;
    Date1 := T.GetValue('PSA_DATEENTREE');
    Date2 := T.GetValue('PSA_DATESORTIE');
    if (Date2 <= idate1900) then Date2 := Date;
    if (Date1 <> idate1900) then
    begin
      if Date1 = Date2 then result := '0'
      else
        Begin
          //PT25 On prends la Date2 +1 jour pour le calcul de durée
          Date2 := Date2 + 1;
          AglNombreDeMoisComplet(Date1, Date2, PremMois, PremAnnee, Duree);
//PT25        If (Date1 = DebutDeMois(Date1)) and (Date2 = FinDeMois(Date2)) then duree := duree + 1 ;  { PT20 }
          result := Duree; //PT12
        End;
    end
    else
      Result := '';
    exit;
  end;

  //DEB PT26
  if FuncName = '@DUREECONTRATDEBCTR' then
  begin
    if T = nil then exit;
    // Récupérer la date de début contrat
    Q := Opensql('SELECT PCI_SALARIE,MAX(PCI_DEBUTCONTRAT) AS DATEDEB,MAX(PCI_ORDRE) AS ORDRE ' +
      'FROM CONTRATTRAVAIL  WHERE PCI_SALARIE="' + CodeSal + '" ' +
      'GROUP BY PCI_SALARIE ', True);
    if not Q.EOF then //PORTAGECWAS
    begin
      DateDeb := Q.FindField('DATEDEB').AsDateTime;
      Ordre := Q.FindField('ORDRE').AsInteger;
    end
    else
    begin
      DateDeb := idate1900;
      Ordre := 0;
    end;
    Ferme(Q);
    if (DateDeb <> idate1900) and (Ordre <> 0) then
      TRech := Tob_Contrat.FindFirst(['PCI_SALARIE', 'PCI_DATEDEB', 'PCI_ORDRE'], [CodeSal, DateDeb, Ordre], False)
    else
      TRech := nil;
    if (TRech <> nil) then
      Date1 := TRech.GetValue('PCI_DEBUTCONTRAT')
    else
      Date1 := idate1900;
    Date2 := T.GetValue('PSA_DATESORTIE');
    if (Date2 <= idate1900) then Date2 := Date;
    if (Date1 <> idate1900) then
    begin
      if Date1 = Date2 then result := '0'
      else
      begin
        Date2 := Date2 + 1;
        AglNombreDeMoisComplet(Date1, Date2, PremMois, PremAnnee, Duree);
        result := Duree;
      end;
    end
    else
      Result := '';
    exit;
  end;
  //FIN PT26

  { DEB PT13 }
  if (FuncName = '@DUREECONTRATJOUR') or (FuncName = '@DUREECONTRATSEMAINE') then  //PT28
  begin
    if T = nil then exit;
    Date1 := T.GetValue('PSA_DATEENTREE');
    Date2 := T.GetValue('PSA_DATESORTIE');
    if (Date2 <= idate1900) then Date2 := Date;  { PT20 18/08/2005 }
    if (Date1 <> idate1900) and (Date2 <> idate1900) then
    begin
      result := Date2 - Date1 + 1;
      //DEB PT28
      if FuncName = '@DUREECONTRATSEMAINE' then
        result := Trunc(result / 7);
      //FIN PT28
    end
    else
      Result := '';
    exit;
  end;
  { FIN  PT13 }

if ((FuncName='@DUREEESSAI') or (FuncName='@DUREEPREAVIS') or
   (FuncName='@DUREEESSAISEMAINE') or (FuncName='@DUREEPREAVISSEMAINE')) or
   (FuncName='@DUREEPREAVISJOUR') then { PT17 } {PT28} {PT31}
  begin
    Duree := 0;
    Q := Opensql('SELECT PCI_SALARIE,MAX(PCI_DEBUTCONTRAT) AS DATEDEB,MAX(PCI_ORDRE) AS ORDRE ' +
      'FROM CONTRATTRAVAIL  WHERE PCI_SALARIE="' + CodeSal + '" ' +
      'GROUP BY PCI_SALARIE ', True);
    if not Q.EOF then //PORTAGECWAS
    begin
      DateDeb := Q.FindField('DATEDEB').AsDateTime;
      Ordre := Q.FindField('ORDRE').AsInteger;
    end
    else
    begin
      DateDeb := idate1900;
      Ordre := 0;
    end;
    Ferme(Q);
    if (DateDeb <> idate1900) and (Ordre <> 0) then
      TRech := Tob_Contrat.FindFirst(['PCI_SALARIE', 'PCI_DATEDEB', 'PCI_ORDRE'], [CodeSal, DateDeb, Ordre], False)
    else
      TRech := nil;
    if (FuncName = '@DUREEESSAI') and (Trech <> nil) then
    begin
      Date1 := TRech.GetValue('PCI_ESSAIDEBUT');
      Date2 := TRech.GetValue('PCI_ESSAIFIN');
      if (Date1 <> idate1900) and (Date2 <> idate1900) then
        AglNombreDeMoisComplet(Date1, Date2, PremMois, PremAnnee, Duree);
      if Date1 = Date2 then result := '0'
      else
        Begin
        If (Date1 = DebutDeMois(Date1)) and (Date2 = FinDeMois(Date2)) then duree := duree + 1 ; { PT20 }
        result := Duree; //PT12
        End;
    end;
    //DEB PT28
    if (FuncName = '@DUREEESSAISEMAINE') and (Trech <> nil) then
    begin
      Date1 := TRech.GetValue('PCI_ESSAIDEBUT');
      Date2 := TRech.GetValue('PCI_ESSAIFIN');
      if Date1 = Date2 then
        result := '0'
      else if (Date1 <> idate1900) and (Date2 <> idate1900) then
      begin
        result := Date2 - Date1 + 1;
        result := Trunc(result / 7);
        exit;
      end;
    end;
    if (FuncName = '@DUREEPREAVISSEMAINE') and (Trech <> nil) then
    begin
      Date1 := TRech.GetValue('PCI_DEBPREAVIS');
      Date2 := TRech.GetValue('PCI_FINPREAVIS');
      if Date1 = Date2 then
        result := '0'
      else if (Date1 <> idate1900) and (Date2 <> idate1900) then
      begin
        result := Date2 - Date1 + 1;
        result := Trunc(result / 7);
        exit;
      end;
    end;
    //FIN PT28
    if (FuncName = '@DUREEPREAVIS') and (Trech <> nil) then { DEB PT17 }
    begin
      Date1 := TRech.GetValue('PCI_DEBPREAVIS');
      Date2 := TRech.GetValue('PCI_FINPREAVIS');
      if (Date1 <> idate1900) and (Date2 <> idate1900) then
        AglNombreDeMoisComplet(Date1, Date2, PremMois, PremAnnee, Duree);
      if Date1 = Date2 then result := '0'
      else
        Begin
        If (Date1 = DebutDeMois(Date1)) and (Date2 = FinDeMois(Date2)) then duree := duree + 1 ; { PT20 }
        result := Duree;
        End;
    end; { FIN PT17 }
    // DEB PT31
    if (FuncName = '@DUREEPREAVISJOUR') and (Trech <> nil) then
      begin
      result := '0';
      Date1 := TRech.GetValue('PCI_DEBPREAVIS');
      Date2 := TRech.GetValue('PCI_FINPREAVIS');
      if (Date1 <> idate1900) and (Date2 <> idate1900) then
      begin
        DiffMoisJour(Date1, Date2, nbMois, nbjour);
        result := inttostr(nbjour + 1);
      end;
      exit;
      end; { FIN PT31}

    result := Duree;
    exit;
  end;

  if FuncName = '@GETLIBELLE' then //DEB PT2-1
  begin
    NomChamp := UpperCase(Params[1]);
    ValChamp := Params[2];
    NomChamp := Trim(Copy(NomChamp, Pos('_', NomChamp), Length(NomChamp)));
    if NomChamp = '' then exit;
    result := PgGetLibelleTablette(NomChamp, ValChamp, Tablette); { PT19 }
    exit;
  end; //FIN PT2-1

  //if T<>nil then T.free;
  { DEB PT18 }
  if FuncName = '@GETORGANISME' then
  begin
    Etab := Params[1];
    Org := Params[2];
    NomChamp := UpperCase(Params[3]);
    Result := '';
    if Assigned(Tob_Organisme) then
    begin
      TRech := Tob_Organisme.FindFirst(['POG_ORGANISME', 'POG_ETABLISSEMENT'], [Org, Etab], False);
      if Assigned(TRech) and (Trech.FieldExists(NomChamp)) then
      begin
        Result := Trech.GetValue(NomChamp);
        if Result = null then Result := '';
      end;
    end;
    exit;
  end;
  { FIN PT18 }

  { DEB PT21 }
    if FuncName = '@GETCUMUL' then
  begin
    StDD := Params[1];
    StDf := Params[2];
    CodeSal := Params[3];
    StCum := Params[4];
    if not IsValidDate(StDD) then begin result := 'Format date début non valide'; exit; end;
    if not IsValidDate(StDF) then begin result := 'Format date fin non valide'; exit; end;
    if StCum = '' then Begin result := 0; exit; End; { PT23 }
    While StCum<>'' do
      Begin
      Cumul := ReadTokenSt(StCum);
      StCumul := StCumul + ' PHC_CUMULPAIE="'+Cumul+'" OR';
      End;
    if StCumul<>'' then StCumul := ' AND ('+Copy(StCumul,1,Length(StCumul)-3)+') ';

    SQL := 'SELECT SUM(PHC_MONTANT) MONT FROM HISTOCUMSAL '+
           'WHERE PHC_DATEDEBUT>="'+USDateTime(StrToDate(StDD))+'"'+
           'AND PHC_DATEFIN<="'+USDateTime(StrToDate(StDF))+'"'+
           'AND PHC_SALARIE="'+CodeSal+'" '+StCumul;
    Q := OpenSql(SQL,True);
    if Not Q.eof then result := Q.FindField('MONT').AsFloat
    else result := 0;
    Ferme(Q);
    exit;
  End;

  if (FuncName = '@GETBASEREM') or (FuncName = '@GETMTREM')  or
     (FuncName = '@GETBASECOT') or (FuncName = '@GETMTSALCOT') or (FuncName = '@GETMTPATCOT') then
  begin
    StDD := Params[1];
    StDf := Params[2];
    CodeSal := Params[3];
    StCum := Params[4];
    if not IsValidDate(StDD) then begin result := 'Format date début non valide'; exit; end;
    if not IsValidDate(StDF) then begin result := 'Format date fin non valide'; exit; end;
    if StCum = '' then Begin result := 0; exit; End;   { PT23 }
    While StCum<>'' do
      Begin
      Cumul := ReadTokenSt(StCum);
      StCumul := StCumul + ' PHB_RUBRIQUE="'+Cumul+'" OR';
      End;
    if StCumul<>'' then StCumul := ' AND ('+Copy(StCumul,1,Length(StCumul)-3)+') ';

    StNat := ' AND PHB_NATURERUB="AAA" ';  StChamp := 'PHB_BASEREM';
    if (FuncName = '@GETMTREM') then  StChamp := 'PHB_MTREM'
    else  if (FuncName = '@GETBASECOT') then  StChamp := 'PHB_BASECOT'
    else  if (FuncName = '@GETMTSALCOT') then  StChamp := 'PHB_MTSALARIAL'
    else  if (FuncName = '@GETMTPATCOT') then  StChamp := 'PHB_MTPATRONAL';
    if (FuncName = '@GETBASECOT') or (FuncName = '@GETMTSALCOT') or (FuncName = '@GETMTPATCOT') then
      StNat := ' AND PHB_NATURERUB="COT" ';
    SQL := 'SELECT SUM('+StChamp+') MONT FROM HISTOBULLETIN '+
           'WHERE PHB_DATEDEBUT>="'+USDateTime(StrToDate(StDD))+'"'+
           'AND PHB_DATEFIN<="'+USDateTime(StrToDate(StDF))+'"'+
           'AND PHB_SALARIE="'+CodeSal+'" '+StCumul + StNat;
    Q := OpenSql(SQL,True);
    if Not Q.eof then result := Q.FindField('MONT').AsFloat
    else result := 0;
    Ferme(Q);
    Exit;
  End;
  { FIN PT21 }

  //DEB PT27
  //Va récupérer la valeur d'un élément national en remontant les différents niveaux possibles
  if FuncName = '@GETELTNAT' then
  begin
    CodeElt := Params[1];
    CodeSal := Params[2];
    Etab := Params[3];
    StDD := Params[4];
    if not IsValidDate(StDD) then
    begin
      result := 'Format date non valide';
      exit;
    end;
    Result := '';

    // Tenir compte du paramsoc pour savoir si on gère les niveaux SAL, POP et ETB
    if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
    begin
      // Niveau SALarié
      TypeNiveau := 'SAL';
      ValeurNiveau := CodeSal;
      Result := RecupValCodeEltDos(CodeElt,CodeSal,TypeNiveau,ValeurNiveau,StDD);

      // Niveau POPulation
{$IFNDEF CPS1}
      if (Result = '') then
      begin
        ValeurNiveau := '';
        Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
          + ' WHERE PNA_SALARIE = "' + CodeSal + '"'
          + ' AND PNA_TYPEPOP = "PAI"', True);  //Anciennement SAL
        if not Q.Eof then
          ValeurNiveau := Q.FindField('PNA_POPULATION').AsString;
        Ferme(Q);
        TypeNiveau := 'POP';
        if (ValeurNiveau <> '') then
          Result := RecupValCodeEltDos(CodeElt,CodeSal,TypeNiveau,ValeurNiveau,StDD);
      end;
{$ENDIF}

      // Niveau ETaBlissement
      if (Result = '') then
      begin
        TypeNiveau := 'ETB';
        ValeurNiveau := Etab;
        Result := RecupValCodeEltDos(CodeElt,CodeSal,TypeNiveau,ValeurNiveau,StDD);
      end;
    end;

    // Niveau DOSsier
    if (Result = '') then
      Result := RecupValCodeEltNat(CodeElt,CodeSal,'DOS',StDD);

    // Niveau STanDard
    if (Result = '') then
      Result := RecupValCodeEltNat(CodeElt,CodeSal,'STD',StDD);

    // Niveau CEGID
    if (Result = '') then
      Result := RecupValCodeEltNat(CodeElt,CodeSal,'CEG',StDD);

    exit;
  end;
  //FIN PT27

end;

//DEB PT27
function TFFusion.RecupValCodeEltDos(CodeElt, CodeSal, TypeNiveau, ValeurNiveau, StDD : String) : String;
var
  Q : TQuery;
  Montant, MontantEuro : double;
begin
  Result := '';
  Montant := 0;
  MontantEuro := 0;

  Q := Opensql('SELECT PED_MONTANT, PED_MONTANTEURO FROM ELTNATIONDOS' +
    ' WHERE PED_CODEELT = "' + CodeElt + '"' +
    ' AND PED_DATEVALIDITE <= "' + USDateTime(StrToDate(StDD)) + '"' +
    ' AND PED_TYPENIVEAU = "' + TypeNiveau + '"' +
    ' AND PED_VALEURNIVEAU = "' + ValeurNiveau + '"' +
    ' ORDER BY PED_DATEVALIDITE DESC', True);
  if not Q.EOF then //PORTAGECWAS
  begin
    Montant := Q.FindField('PED_MONTANT').AsFloat;
    MontantEuro := Q.FindField('PED_MONTANTEURO').AsFloat;
    Result := DetermineAlsaceOuNon(CodeSal,Montant,MontantEuro);
  end;
  Ferme(Q);
end;

function TFFusion.DetermineAlsaceOuNon(CodeSal : String; Montant, MontantEuro : Double): String;
var
  Q : TQuery;
begin
  Result := '';
  // Rechercher si l'établissement du salarié est Alsace ou non
  Q := OpenSQL('SELECT ETB_REGIMEALSACE FROM ETABCOMPL'
    + ' LEFT JOIN SALARIES ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT'
    + ' WHERE PSA_SALARIE = "' + CodeSal + '"',True);
  if not Q.Eof then
  begin
    if (Q.FindField('ETB_REGIMEALSACE').AsString = 'X') then
    begin
      if (Montant <> 0) then
        Result := FloatToStr(Montant);
    end
    else
    begin
      if (MontantEuro <> 0) then
        Result := FloatToStr(MontantEuro);
    end;
  end
  else
  begin
    if (MontantEuro <> 0) then
      Result := FloatToStr(MontantEuro);
  end;
  Ferme(Q);
end;

function TFFusion.RecupValCodeEltNat(CodeElt, CodeSal, Predefini, StDD: String): String;
var
  Q : TQuery;
  Montant, MontantEuro : double;
  NoDossier, Convention : String;
begin
  Result := '';
  Montant := 0;
  MontantEuro := 0;

  if (Predefini = 'DOS') then
    NoDossier := PgRendNoDossier()
  else
    NoDossier := '000000';

  if (Predefini = 'STD') then
  begin
    Q := OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE = "' + CodeSal + '"',True);
    if not Q.Eof then
      Convention := Q.FindField('PSA_CONVENTION').AsString
    else
      Convention := '000';
    Ferme(Q);
  end
  else
    Convention := '000';

  Q := Opensql('SELECT PEL_MONTANT, PEL_MONTANTEURO FROM ELTNATIONAUX' +
    ' WHERE PEL_CODEELT = "' + CodeElt + '"' +
    ' AND PEL_DATEVALIDITE <= "' + USDateTime(StrToDate(StDD)) + '"' +
    ' AND PEL_PREDEFINI = "' + Predefini + '"' +
    ' AND PEL_NODOSSIER = "' + NoDossier + '"' +
    ' AND PEL_CONVENTION = "' + Convention + '"' +
    ' ORDER BY PEL_DATEVALIDITE DESC', True);
  if not Q.EOF then //PORTAGECWAS
  begin
    Montant := Q.FindField('PEL_MONTANT').AsFloat;
    MontantEuro := Q.FindField('PEL_MONTANTEURO').AsFloat;
    Result := DetermineAlsaceOuNon(CodeSal,Montant,MontantEuro);
  end;
  Ferme(Q);

  if (Result = '') and (Predefini = 'STD') and (Convention <> '000') then
  begin
    Convention := '000';
    Q := Opensql('SELECT PEL_MONTANT, PEL_MONTANTEURO FROM ELTNATIONAUX' +
      ' WHERE PEL_CODEELT = "' + CodeElt + '"' +
      ' AND PEL_DATEVALIDITE <= "' + USDateTime(StrToDate(StDD)) + '"' +
      ' AND PEL_PREDEFINI = "' + Predefini + '"' +
      ' AND PEL_NODOSSIER = "' + NoDossier + '"' +
      ' AND PEL_CONVENTION = "' + Convention + '"' +
      ' ORDER BY PEL_DATEVALIDITE DESC', True);
    if not Q.EOF then //PORTAGECWAS
    begin
      Montant := Q.FindField('PEL_MONTANT').AsFloat;
      MontantEuro := Q.FindField('PEL_MONTANTEURO').AsFloat;
      Result := DetermineAlsaceOuNon(CodeSal,Montant,MontantEuro);
    end;
    Ferme(Q);
  end;

end;
//FIN PT27

function TFFusion.GetCalendrier(VSalarie: string): string;
var
  Trech, Tsal: Tob;
  Champ, ValChamp, Semaine, StJr, LibJour, HDeb1, HFin1, HDeb2, HFin2: string;
  Calend, Etab, StandCalend: string;
  Q: TQuery;
  Jour, Jr: Integer;
begin
  Jour := 1;
  Semaine := '';
  Calend := '';
  ValChamp := ''; //PORTAGECWAS
  TSal := Tob_Salaries.FindFirst(['PSA_SALARIE'], [VSalarie], False);
  if TSal <> nil then
  begin
    Calend := TSal.getvalue('PSA_CALENDRIER');
    StandCalend := TSal.getvalue('PSA_STANDCALEND');
    Etab := TSal.getvalue('PSA_ETABLISSEMENT');
  end;

  if StandCalend = 'PER' then //1 - on cherche au niveau calendrier spécifique salarie
  begin
    Tob_Calendrier.Detail.Sort('ACA_SALARIE;ACA_JOUR');
    Champ := 'ACA_SALARIE';
    ValChamp := VSalarie;
  end;
  if (StandCalend = 'ETS') and (Calend <> '') then //2 - on cherche au niveau standard de calendrier salarie
  begin
    Tob_Calendrier.Detail.Sort('ACA_STANDCALEN;ACA_JOUR');
    Champ := 'ACA_STANDCALEN';
    ValChamp := Calend;
  end;
  if (StandCalend = 'ETB') and (etab <> '') then //3 - on cherche au niveau calendrier standard établissement
  begin
    Q := OpenSQL('SELECT ETB_STANDCALEND FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' + Etab + '"', TRUE);
    if not Q.EOF then //PORTAGECWAS
      Calend := Q.FindField('ETB_STANDCALEND').asstring;
    Ferme(Q);
    Champ := 'ACA_STANDCALEN';
    ValChamp := Calend;
  end;

  if (Tob_Calendrier <> nil) and (valchamp <> '') then //PORTAGECWAS
    TRech := Tob_Calendrier.FindFirst([Champ, 'ACA_JOUR'], [ValChamp, Jour], False)
  else TRech := nil;
  while TRech <> nil do
  begin
    if (TRech.GetValue('ACA_DUREE') <> 0) then
    begin
      Jr := Jour;
      if Jr = 7 then Jr := 1 else Jr := Jr + 1;
      if Jr = 1 then StJr := '1DI' else
        if Jr = 2 then StJr := '2LU' else
        if Jr = 3 then StJr := '3MA' else
        if Jr = 4 then StJr := '4ME' else
        if Jr = 5 then StJr := '5JE' else
        if Jr = 6 then StJr := '6VE' else
        if Jr = 7 then StJr := '7SA';
      LibJour := RechDom('YYJOURSSEMAINE', StJr, False) + ' ';
      HDeb1 := FloatToStr(TRech.GetValue('ACA_HEUREDEB1'));
      if HDeb1 <> '0' then HDeb1 := HDeb1 + 'h - ';
      HFin1 := FloatToStr(TRech.GetValue('ACA_HEUREFIN1'));
      if HDeb1 <> '0' then HFin1 := HFin1 + 'h ';
      HDeb2 := FloatToStr(TRech.GetValue('ACA_HEUREDEB2'));
      if HDeb2 <> '0' then    //PT29
        HDeb2 := HDeb2 + 'h - '
      else
        HDeb2 := '';          //PT29
      HFin2 := FloatToStr(TRech.GetValue('ACA_HEUREFIN2'));
      if HFin2 <> '0' then    //PT29
        HFin2 := HFin2 + 'h '
      else                    //PT29
        HFin2 := '';
      Semaine := Semaine + ' ' + LibJour + HDeb1 + HFin1 + HDeb2 + HFin2;
    end;
    Jour := Jour + 1;
    TRech := Tob_Calendrier.FindNext([Champ, 'ACA_JOUR'], [ValChamp, Jour], False);
  end;
  result := Semaine;
  //TSal.free; TSal:=nil;
  //Trech.free;TRech:=nil;
end;

function TFFusion.GetCodeSalarie(iIndx: integer): string;
var
  Sal: string;
begin
  Result := '';
  if ((iIndx - 1) < 0) then exit;
  if Argument = nil then exit;
  if Argument.count <= 0 then exit;
  Sal := Argument.Strings[iIndx - 1]; //IntToStr(GetPositionSal(VarIndx));
  if (isnumeric(Sal) and (VH_PAIE.PgTypeNumSal = 'NUM')) then Sal := ColleZeroDevant(StrToInt(Sal), 10);
  result := Sal;
end;
{
function TFFusion.GetTablettelibelle(Champ,value: String): String;
Var
ipos : integer;
Q : TQuery;
Libelle,Suffixe,TypeChamp : string;
begin
result:=value;
if (Champ='PSA_SALARIE')  then exit;
//Test si champ de type combo ou varchar(17)
Q:=OpenSql('SELECT DH_TYPECHAMP FROM DECHAMPS WHERE DH_NOMCHAMP="'+Champ+'"',True);
if not Q.eof then //PORTAGECWAS
  TypeChamp:=Q.FindField('DH_TYPECHAMP').asstring
else
  TypeChamp:='';
Ferme(Q);
if (Pos('VARCHAR',TypeChamp)=0) and (TypeChamp<>'COMBO') then exit;
if (TypeChamp='VARCHAR(35)') and (Champ<>'PSA_LIBELLEEMPLOI')  then exit;
Suffixe:=''; libelle:='';
iPos:=Pos('_',Champ);
if ipos>0 then Suffixe:=Trim(Copy(Champ,Ipos+1,Length(Champ)));
//Recherche Tablette de la paie
if Suffixe='' then begin result:=value; exit; end;
Q:=OpenSql('SELECT DO_COMBO FROM DECOMBOS '+
'WHERE DO_DOMAINE="P" AND DO_NOMCHAMP LIKE "'+suffixe+'" '+
'ORDER BY DO_DOMAINE',True);
if not Q.Eof then //PORTAGECWAS
  Libelle:=RechDom(Q.FindField('DO_COMBO').asstring,value,False);
Ferme(Q);
//Recherche Autre tablette COMMUNE (TT,YY)
if Libelle='' then
  Begin
  Q:=OpenSql('SELECT DO_COMBO FROM DECOMBOS '+
  'WHERE DO_DOMAINE<>"P" AND DO_NOMCHAMP LIKE "'+suffixe+'" '+
  'ORDER BY DO_DOMAINE',True);
  While not Q.eof do
    Begin
    if libelle<>'' then Break;
    Libelle:=RechDom(Q.FindField('DO_COMBO').asstring,value,False);
    Q.next;
    End;
  Ferme(Q);
  End;
if libelle<>'' then  result:=Libelle;
end;
}

procedure TFFusion.OnGetList(Sender: TObject; IdentList: string;
  var List: TStringList);
var
  St: string;
  i: integer;
  T: Tob;
begin
  { DEB PT14 }
  if IdentList = 'PAY_DECLARANT' then
  begin
    i := 0;
    if Assigned(Tob_Declarant) then
    begin
      for i := 0 to Tob_Declarant.Detail.count - 1 do
      begin
        Tob_Declarant.Detail[i].AddChampSupValeur('INDEX', i + 1);
        Tob_Declarant.Detail[i].AddChampSupValeur('ORDRE', i + 1);
        Tob_Declarant.Detail[i].AddChampSupValeur('PDA_DENOMINATION', Tob_Declarant.Detail[i].GetValue('PDA_LIBELLE') + ' ' + Tob_Declarant.Detail[i].GetValue('PDA_PRENOM'));
        St := RechDom('YYCIVILITE', Tob_Declarant.Detail[i].GetValue('PDA_CIVILITE'), False);
        St := St + ' ' + Tob_Declarant.Detail[i].GetValue('PDA_LIBELLE') + ' ' + Tob_Declarant.Detail[i].GetValue('PDA_PRENOM');
        if Tob_Declarant.Detail[i].GetValue('PDA_QUALDECLARANT') = 'AUT' then
          Tob_Declarant.Detail[i].PutValue('PDA_QUALDECLARANT', Tob_Declarant.Detail[i].GetValue('PDA_AUTRE'))
        else
          Tob_Declarant.Detail[i].PutValue('PDA_QUALDECLARANT', RechDom('PGQUALDECLARANT2', Tob_Declarant.Detail[i].GetValue('PDA_QUALDECLARANT'), False));
        St := St + ' - ' + Tob_Declarant.Detail[i].GetValue('PDA_QUALDECLARANT');
        List.Add(St);
      end;
      i := Tob_Declarant.Detail.count;
    end
    else
      Tob_Declarant := TOB.Create('Declarant attestation', nil, -1);
    { Ajout de l'option autre }
    if Assigned(Tob_Declarant) then
    begin
      T := TOB.Create('Declarant attestation', Tob_Declarant, -1);
      T.AddChampSupValeur('ORDRE', i + 1);
      T.AddChampSupValeur('INDEX', -1);
      T.AddChampSupValeur('PDA_DENOMINATION', 'Autre');
      List.Add('Autre');
    end;
  end;
  { FIN PT14 }
end;

procedure TFFusion.OnInitialization(Sender: TObject);
var
  Q: TQuery;
  St: string;
begin
  { DEB PT14 }
  if Objet = 'CERTIFICAT' then St := 'CER'
  else if Objet = 'SOLDE' then St := 'SLD'
  else if Objet = 'CONTRAT' then St := 'CON'
  else St := '';
  Q := OpenSql('SELECT * FROM DECLARANTATTEST WHERE (PDA_TYPEATTEST="" OR PDA_TYPEATTEST LIKE "%' + St + '%")', True);
  if not Q.Eof then
  begin
    Tob_Declarant := TOB.Create('Declarant attestation', nil, -1);
    Tob_Declarant.LoadDetailDB('Declarant attestation', '', '', Q, False);
  end
  else
    FreeAndNil(Tob_Declarant);
  Ferme(Q);
  { FIN PT14 }
end;

procedure TFFusion.OnSetVar(Sender: TObject; VarName: string;
  VarIndx: Integer; Value: variant);
begin
  ;
end;

procedure ValidateScript(theScript: string; var bValid: boolean);
begin
  ;
end;



end.

