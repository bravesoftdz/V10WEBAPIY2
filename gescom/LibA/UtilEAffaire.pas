unit UtilEAffaire;

interface
uses UTOB,UTOM,AffaireUtil,CalcOleGenericAff,
  HCtrls,HEnt1,EntGC,Ent1,UtilPGI{$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF};

Function TraiteUneEAffaire ( TobEAff : TOB) : Integer;
Function CompleteAffaire ( TobAff : TOB) : integer;
procedure EAffaireVersAffaire(TobEAff,TobAff : TOB);

implementation

Function TraiteUneEAffaire ( TobEAff : TOB) : Integer;
Var TobAff : TOB;
    TOM_AFF : TOM;
    //NumErr : integer;
    bok : Boolean;
begin
  TobAff := TOB.Create ('AFFAIRE', Nil,-1);
  TOM_AFF := CreateTom('AFFAIRE',Nil,False, False);
  try
    // Passage par OnNewRecord de la TOM
    TOM_AFF.InitTob(TobAff);

    EAffaireVersAffaire(TobEAff,TobAff);
    Result := CompleteAffaire(TobAff);
  // mcd 06/12/02 ne pas positioner V_PgiErro. ce n'est pa sune err, mais une erreur fct, à traiter par nous  If Result <> 0 then Begin V_PGI.IoError:=oeUnknown; Exit; End;
    If Result <> 0 then Begin  Exit; End;

    // Vérification passage par OnUpdateRecord de la TOM
    TobAff.AddChampSupValeur ('CreationAuto','X');     // Test dans le OnUpdate de la TOM Affaire
    TobAff.AddChampSupValeur ('CalculEcheances','X');
    bOk := TOM_AFF.VerifTOB (TobAff);
   // mcd 06/12/02 If not (bOK) then Begin V_PGI.IoError:=oeUnknown; Exit; Result:=9;  End;
    If not (bOK) then Begin V_PGI.IoError:=oeUnknown; Result:=9; Exit;  End;
    TobAff.InsertDB(Nil);
  Finally
    TOM_AFF.Free;
    TobAff.Free;
  end;
end;

procedure EAffaireVersAffaire(TobEAff,TobAff : TOB);
Var Suffixe : String;
    i : integer;
begin
  for i:= 1 to TobEAff.NbChamps do
  begin
    Suffixe := TobEAff.GetNomChamp(i);
    Delete(Suffixe, 1, Pos('_', Suffixe)-1);
    // gm ???? il faut vérifier sur la tob  source  (tobeaff) ???
    if TobAff.FieldExists('AFF'+Suffixe) then
      TobAff.PutValue('AFF'+Suffixe, TobEAff.GetValue('EAF'+Suffixe));
  end;
end;

Function CompleteAffaire ( TobAff : TOB) : integer;
Var bOK : Boolean;
    CodeAff, Aff0, Aff1, Aff2, Aff3, Aff4,Periode : string;
begin
bOK := True; Result := 0;
// Code affaire ( Affaire 1,2,3 toujours aliemntée sauf si compteur auto)
if TobAff.GetValue ('AFF_AFFAIRE0')= ''  then TobAff.PutValue('AFF_AFFAIRE0','A');
if TobAff.GetValue ('AFF_AVENANT')= ''   then TobAff.PutValue('AFF_AVENANT','00');
Aff0 := TobAff.GetValue('AFF_AFFAIRE0'); Aff1 := TobAff.GetValue('AFF_AFFAIRE1');
Aff2 := TobAff.GetValue('AFF_AFFAIRE2'); Aff3 := TobAff.GetValue('AFF_AFFAIRE3');
Aff4 := TobAff.GetValue('AFF_AVENANT');

{if ((VH_GC.AFFormatExer)  <>'AUC') and (aff3='') then
        begin   // si gestion exercice, on regarde si mission pour tiers existe déjà  (si compteur= Affaire3 pas renseigné)
                // pas mis en place, car dans ce cas on a un réseau surchargé ... vient du fait que l'on est dans une transaction ?? pas trouvé au 19/03/02
        champ :=  'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_TIERS = "'
              + TobAff.GetValue('AFF_TIERS') +'" AND AFF_AFFAIRE1 = "'+ aff1 +'"  and AFF_AFFAIRE0="' + aff0+'" '  ;
        if VH_GC.CleAffaire.NbPartie =3 then champ := champ + 'AND AFF_AFFAIRE2 = "'+ aff2+'"';
        if  ExisteSQL (champ) then
            Begin  // cas ou mission existe déjà pour ce client, on refuse la création
            result:=1;
            exit;
            end ;
   end;  }
CodeAff := CodeAffaireRegroupe(Aff0, Aff1, Aff2, Aff3, Aff4,taCreat,True,False,True);
if CodeAff = '' then begin Result:=1; Exit; end;
if ExisteAffaire (CodeAff,'') then begin Result:=1; Exit; end
                              else TobAff.PutValue('AFF_AFFAIRE0',Copy(CodeAff,1,1));
               // MCD 20/06/02 plante agl 540               else TobAff.PutValue('AFF_AFFAIRE0',CodeAff);

TobAff.PutValue('AFF_AFFAIRE1', Aff1);  TobAff.PutValue('AFF_AFFAIRE2', Aff2);
TobAff.PutValue('AFF_AFFAIRE3',Aff3);
TobAff.PutValue('AFF_AFFAIRE',CodeAff); TobAff.PutValue('AFF_AFFAIREREF',CodeAff);
// *** Statut d'affaires ***
if TobAff.GetValue ('AFF_AFFAIRE0')= 'P' then TobAff.PutValue('AFF_STATUTAFFAIRE', 'PRO')
                                         else TobAff.PutValue('AFF_STATUTAFFAIRE', 'AFF');
// *** Etat ***
if TobAff.GetValue ('AFF_ETATAFFAIRE')= ''  then
   TobAff.PutValue('AFF_ETATAFFAIRE','ENC')
else
   bOK := ExisteSQL ('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="AET" AND CC_CODE="'+TobAff.GetValue('AFF_ETATAFFAIRE')+'"');
if Not (bOk) then begin Result:=3; Exit; end;
// *** libellé ***
if TobAff.GetValue ('AFF_LIBELLE')= ''  then TobAff.PutValue ('AFF_LIBELLE','Affaire : '+ CodeAffaireAffiche(CodeAff));
// *** Client ***
if TobAff.GetValue ('AFF_TIERS')= ''  then begin Result :=7; Exit; end;
if (IsTiersFerme(TobAff.GetValue ('AFF_TIERS'))) then begin Result :=6; Exit; end;
if Not( ExisteSQL ('SELECT T_TIERS From TIERS WHERE T_TIERS="'+ TobAff.GetValue ('AFF_TIERS')+ '" AND '+ FabriqueWhereNatureAuxiAff(TobAff.GetValue('AFF_STATUTAFFAIRE')))) then
   begin Result :=4; Exit; end;
// *** devise ***
if TobAff.GetValue('AFF_DEVISE')='' then TobAff.PutValue('AFF_DEVISE',V_PGI.DevisePivot);
   bOK := ExisteSQL ('SELECT D_DEVISE FROM DEVISE WHERE D_DEVISE="'+TobAff.GetValue('AFF_DEVISE')+'"');
if Not (bOk) then begin Result:=8; Exit; end;

// controle d'init sur les champs ayant des valeurs par défaut.
if TobAff.GetValue('AFF_GENERAUTO')  = '' then TobAff.PutValue('AFF_GENERAUTO',VH_GC.AFFGenerAuto);
if TobAff.GetValue('AFF_PERIODICITE')= '' then TobAff.PutValue('AFF_PERIODICITE','M');
if TobAff.GetValue('AFF_ETABLISSEMENT')  = '' then TobAff.PutValue('AFF_ETABLISSEMENT',VH^.EtablisDefaut);
if TobAff.GetValue('AFF_REPRISEACTIV')  = '' then
   begin
   if (TobAff.GetValue('AFF_GENERAUTO') <> 'ACT') then TobAff.PutValue('AFF_REPRISEACTIV',VH_GC.AFRepriseActiv)
                                                  else TobAff.PutValue('AFF_REPRISEACTIV','TOU');
   end;
Periode := TobAff.GetValue ('AFF_PERIODICITE');
if (Periode <> 'A') And (Periode <> 'M') And (Periode <> 'J') then
   TobAff.PutValue('AFF_PERIODICITE','M');
// controle les booleans
if TobAff.GetValue('AFF_SAISIECONTRE') = '' then TobAff.PutValue('AFF_SAISIECONTRE','-');
if TobAff.GetValue('AFF_CALCTOTHTGLO') = '' then TobAff.PutValue('AFF_CALCTOTHTGLO','-');
// interval
if TobAff.GetValue('AFF_INTERVALGENER') = 0 then TobAff.PutValue('AFF_INTERVALGENER',1);
// contrôle des dates  ???
if TobAff.GetValue('AFF_DATEFINGENER') = iDate1900 then TobAff.PutValue('AFF_DATEFINGENER',iDate2099);
if TobAff.GetValue('AFF_DATEFIN') = iDate1900 then TobAff.PutValue('AFF_DATEFIN',iDate2099);

// forcer pour ALGOE, mettre acces au client pour avoir la bonne valeur
// nb : le loadtobfrom file charge les boolean à '-'

  TobAff.PutValue('AFF_AFFAIREHT','X');
end;

end.
