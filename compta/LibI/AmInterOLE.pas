{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/01/2005
Modifié le ... : 11/01/2005
Description .. : - FQ 15239 - CA - 11/01/2005 - On ne prend plus en
Suite ........ : compte les
Suite ........ : immobilisations de l'historique dans les calculs.
Suite ........ : 28/02/2006 - MVG - Report des modifs de C.AYEL
Mots clefs ... :

FQ 17756 : TGA 31/03/2006 test sur composant pour les regroupements
FQ 18123 : MVG 16/06/2006 Mutation dans la même tranche de liasse
- BTY 11/06 FQ 18264 : Ouvrir la taxe professionnelle dans Publifi aux Locations longue durée
- BTY 11/06 FQ 19036 : Pour la taxe pro, nouvelle sélection sur les immos
FQ 19514 : MVG 15/01/2007 Alimentation type de dérogatoire
FQ 17512 : MBO   /03/2007 modif pour dérogatoire (chantier fiscal)
FQ 20256 : BTY   05/07 En nrécup dérogatoire, filtrer les immos sur i_typederogatoire
FQ 21159 : MBO 23/07/2007 : duplication des verbes en filtrant sur l'établissement
- MBO 20/09/07 Report de modif ds f° AmGetLesImmosEtab (modif C.Malié)
- MBO 03/10/07 Report de modif ds f° Etab (modif C.Malié) identifées CMX 01/10/2007
FQ 21754 - MBO 29/10/07 Modif appel planinfo.Calcul pour nveaux paramètres liés au calcul de cession (tjrs à false ici)
*****************************************************************}
unit AmInterOLE;

interface

uses  SysUtils,
{$IFDEF eAGLClient}
{$ELSE}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF eAGLClient}
      utob,
      Ent1,
      Hent1,
      hctrls,
      implaninfo,
      {$IFDEF MODENT1}
      CPTypeCons,
      {$ENDIF MODENT1}
      Imoutgen;

function AmGetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup : string; Eligible : string = '') : Variant ;
function AmGetAugAcqMut(CpteInf,CpteSup : string) : Variant ;
function AmGetDiminutionMutation(CpteInf,CpteSup : string) : Variant ;
function AmGetDotationExo(CpteInf,CpteSup, Methode : string) : Variant ;
function AmGetDiminutionAmort(CpteInf,CpteSup : string) : Variant ;
function AmGetListeLieuImmo : Variant ;
function AmGetDiminutionSortie(CpteInf,CpteSup : string) : Variant ;
function AmGetValHTImmo ( pstCpteInf, pstCpteSup : string; pstParam : string ) : variant;
function AmGetDerogatoire(CpteInf,CpteSup, Methode : string) : Variant ;
function AmGetLesImmos(stType: string; iVersion: integer; laTob: TOB): TOB;

// fq 21159
// nouvelles fonctions par etablissement mbo 20.07.07
// cmx - 01/10/07 - ajout du paramètre Where ds toutes ces f°
function AmGetAugAcqMutEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
function AmGetDiminutionMutationEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
function AmGetDotationExoEtab(CpteInf,CpteSup, Methode, Etablissement : string; Where: string = '') : Variant ;
function AmGetDiminutionAmortEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
function AmGetDiminutionSortieEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
function AmGetValHTImmoEtab ( pstCpteInf, pstCpteSup : string; pstParam, Etablissement : string; Where: string = '') : variant;
function AmGetDerogatoireEtab(CpteInf,CpteSup, Methode, Etablissement : string; Where: string = '') : Variant ;
function AmGetLesImmosEtab(stType: string; iVersion: integer; laTob: TOB; Etablissement: string; Where: string = ''): TOB;


implementation

procedure AmBourreBorneCompte ( var St : String ; Bourre : Char );
var Lg,ll,i : Integer ;
begin
  Lg:=VH^.Cpta[fbGene].Lg ;
  If Length(St)>Lg Then
  begin
    St:=Trim(Copy(St,1,Lg)) ;
    exit;
  end;
  ll:=Length(St) ;
  If ll<Lg then
  BEGIN
    for i:=ll+1 to Lg do St:=St+Bourre ;
  END;
end ;

procedure AmBourreBornes (var St1,St2 : string);
begin
  AmBourreBorneCompte ( St1, '0');
  AmBourreBorneCompte ( St2, '9');
end;

{ MBO FQ 18185 22/05/2006
function AmGetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup : string; Eligible : string = '') : Variant ;
var Q : TQuery;
    sSelect : string;
    sWhereCompte : string;
    sWhereBien : string;
    sWhereLieu : string;
    sWhereEligible : string;
begin
  AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT SUM(I_BASETAXEPRO) FROM IMMO';
  sWhereCompte := ' WHERE I_ETAT<>"FER" AND I_COMPTEREF>="'+CpteInf+'" AND I_COMPTEREF <="'+CpteSup+'"';
  if NatureBien <> '' then sWhereBien := ' AND I_NATUREBIEN="'+NatureBien+'"'
                      else sWhereBien := '';
  sWhereLieu := '';
  if LieuInf <> '' then sWhereLieu := ' AND I_LIEUGEO>="'+LieuInf+'"';
  if LieuSup <> '' then sWhereLieu := sWhereLieu+' AND I_LIEUGEO<="'+LieuSup+'"';
  sWhereEligible := '';
  if Eligible = '-' then sWhereEligible := ' AND (I_TABLE9="-" OR I_TABLE9="") '
  else if Eligible = 'X' then sWhereEligible := ' AND I_TABLE9="X"';
  Q:=OpenSQL (sSelect+sWhereCompte+sWhereBien+sWhereLieu+sWhereEligible,True);
  if not Q.Eof then Result:=Arrondi(Q.Fields[0].AsFloat,V_PGI.OkDecV) ;
  Ferme(Q);
end ;
}
function AmGetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup : string; Eligible : string = '') : Variant ;
var Q : TQuery;
    sSelect : string;
    sWhereCompte : string;
    sWhereBien : string;
    sWhereLieu : string;
    sWhereEligible : string;
begin
  AmBourreBornes (CpteInf,CpteSup);

  // FQ 19036
  {sSelect := 'SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,I_BASETAXEPRO,I_COMPTEIMMO,I_NATUREIMMO, I_NATUREBIEN FROM IMMOLOG LEFT JOIN IMMO ON (I_IMMO=IL_IMMO)';
  // FQ 18264
  //sWhereCompte := ' WHERE I_ETAT<>"FER" AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="CB")'+
  sWhereCompte := ' WHERE I_ETAT<>"FER" AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC")'+
        ' AND (IL_TYPEOP="ACQ") AND (IL_DATEOP<="' +
        USDateTime(VH^.Encours.Fin)
        +'") AND I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'"'; }

  sSelect := 'SELECT I_IMMO,I_BASETAXEPRO,I_COMPTEIMMO,I_NATUREIMMO,I_NATUREBIEN,I_DATEFINCB,I_OPELEVEEOPTION FROM IMMO';
  sSelect := sSelect +
             ' LEFT JOIN IMMOLOG IL1 ON (I_IMMO=IL1.IL_IMMO AND IL1.IL_TYPEOP="ACQ")' +
             ' LEFT JOIN IMMOLOG IL2 ON (I_IMMO=IL2.IL_IMMO AND IL2.IL_TYPEOP="CES")' +
             ' LEFT JOIN IMMOLOG IL3 ON (I_IMMO=IL3.IL_IMMO AND IL3.IL_TYPEOP="LEV")';
  sWhereCompte := ' WHERE I_ETAT<>"FER" AND I_NATUREIMMO<>"FI"'+
                  ' AND (IL1.IL_DATEOP<="' + USDateTime(VH^.Encours.Fin) +  '")' +
                  ' AND ((IL2.IL_DATEOP>="' + USDateTime(VH^.Encours.Fin) +  '")' +
                     ' OR IL2.IL_DATEOP IS NULL)' +
                  ' AND ((I_OPELEVEEOPTION="X" AND IL3.IL_DATEOP>"'+USDateTime(VH^.Encours.Fin)+'")'+
                     ' OR (I_OPELEVEEOPTION<>"X"))' +
                  ' AND (((I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND (I_DATEFINCB>="' +
                          USDateTime(VH^.Encours.Fin)+'"))' +
                     ' OR (I_NATUREIMMO<>"CB" AND I_NATUREIMMO<>"LOC"))' +
                  ' AND I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'"';
  if NatureBien <> '' then sWhereBien := ' AND I_NATUREBIEN="'+NatureBien+'"'
                      else sWhereBien := '';
  sWhereLieu := '';
  if LieuInf <> '' then sWhereLieu := ' AND I_LIEUGEO>="'+LieuInf+'"';
  if LieuSup <> '' then sWhereLieu := sWhereLieu+' AND I_LIEUGEO<="'+LieuSup+'"';
  sWhereEligible := '';
  if Eligible = '-' then sWhereEligible := ' AND (I_TABLE9="-" OR I_TABLE9="") '
  else if Eligible = 'X' then sWhereEligible := ' AND I_TABLE9="X"';

  Q:=OpenSQL (sSelect+sWhereCompte+sWhereBien+sWhereLieu+sWhereEligible,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    result :=  result + Q.FindField('I_BASETAXEPRO').AsFloat;

  Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;


Function AmGetAugAcqMut(CpteInf,CpteSup : string) : Variant ;
var Q : TQuery;
    sSelect, sWhere : string;
    CodeEnCours : string;
begin
  AmBourreBornes (CpteInf,CpteSup);
  CodeEnCours:='' ;
  sSelect:='SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,I_MONTANTHT,I_VALEURACHAT,I_COMPTEREF FROM IMMOLOG LEFT JOIN IMMO ON (I_IMMO=IL_IMMO)';
  sWhere:=' WHERE I_ETAT<>"FER" AND (IL_TYPEOP="ACQ" OR IL_TYPEOP="MUT") AND (IL_DATEOP>="'+
        USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
        +'") AND (I_COMPTEREF>="'+CpteInf+'" AND I_COMPTEREF <="'+CpteSup+'") ORDER BY IL_IMMO,IL_DATEOP DESC';
  Q:=OpenSQL(sSelect+sWhere,True) ;
  Result:=0 ;
  while not Q.Eof do
  begin
    if Q.FindField('IL_IMMO').AsString <> CodeEnCours then
    begin
//    result :=  result + Q.FindField('I_MONTANTHT').AsFloat;
      if Q.FindField('IL_TYPEOP').AsString = 'ACQ' then
        result :=  result + Q.FindField('I_VALEURACHAT').AsFloat
      else result := result + Q.FindField('I_MONTANTHT').AsFloat;
      CodeEnCours := Q.FindField('IL_IMMO').AsString;
    end;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionSortie(CpteInf,CpteSup : string) : Variant ;
var Q : TQuery;
    sSelect, sWhere : string;
begin
  AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT IL_VOCEDEE,IL_TYPEOP,IL_DATEOP,I_COMPTEREF FROM IMMOLOG LEFT JOIN IMMO ON (IL_IMMO=I_IMMO)';
  sWhere := ' WHERE I_ETAT<>"FER" AND IL_TYPEOP LIKE "CE%" AND (IL_DATEOP>="'+
        USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
         +'") AND (I_COMPTEREF>="'+CpteInf+'" AND I_COMPTEREF <="'+CpteSup+'")';
  Result:=0 ;
  Q:=OpenSQL (sSelect+sWhere,True);
  while not Q.Eof do
  begin
    Result:=Result+Q.FindField('IL_VOCEDEE').AsFloat ;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionMutation(CpteInf,CpteSup : string) : Variant ;
var Q : TQuery;
    sSelect, sWhere, CodeEnCours : string;
begin
  AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,IL_CPTEMUTATION,I_COMPTEREF,I_MONTANTHT FROM IMMOLOG LEFT JOIN IMMO ON (IL_IMMO=I_IMMO)';
  sWhere := ' WHERE I_ETAT<>"FER" AND IL_TYPEOP="MUT" AND (IL_DATEOP>="'+
          USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
          +'") AND (IL_CPTEMUTATION>="'+CpteInf+'" AND IL_CPTEMUTATION <="'+CpteSup+'")'
          +'  ORDER BY IL_IMMO,IL_DATEOP DESC';
// MVG FQ 18123 +' AND (I_COMPTEREF<"'+CpteInf+'" OR I_COMPTEREF >"'+CpteSup+'") ORDER BY IL_IMMO,IL_DATEOP DESC';
  Q:=OpenSQL(sSelect+sWhere,True) ;
  Result:=0 ; CodeEnCours:='' ;
  while not Q.Eof do
  begin
    if CodeEnCours <> Q.FindField('IL_IMMO').AsString then
    begin
      Result:=Result+Q.FindField('I_MONTANTHT').AsFloat ;
      CodeEnCours:=Q.FindField('IL_IMMO').AsString ;
    end;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDotationExo(CpteInf,CpteSup, Methode : string) : Variant ;
var Q : TQuery;
    Q2: TQuery;
    sSelect, sWhere , sWhereCompte : string;
    sSelect2, sWhere2 : string;
    mDerog : double;
begin
  AmBourreBornes(CpteInf,CpteSup);
  sSelect:='Select IA_MONTANTECO,IA_MONTANTFISCAL,IA_CESSIONECO,IA_CESSIONFISCAL,I_PLANACTIF,I_METHODEECO,I_METHODEFISC,I_MONTANTEXC,I_MONTANTEXCCED, I_OPECESSION FROM IMMOAMOR LEFT JOIN IMMO ON (I_IMMO=IA_IMMO)' ;
  // CA - 04/09/2002 - On ne prend en compte que les immos en pleine propriété
  sWhere:='WHERE I_ETAT<>"FER" AND I_NATUREIMMO="PRO" AND (IA_NUMEROSEQ=I_PLANACTIF AND IA_DATE="'+UsDateTime(VH^.Encours.Fin)+'")' ;
  sWhereCompte:=' AND (IA_COMPTEIMMO>="'+CpteInf+'" AND IA_COMPTEIMMO <="'+CpteSup+'")' ;
  Q:=OpenSQL (sSelect+sWhere+sWhereCompte,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    if (Q.FindField('I_METHODEFISC').AsString <> '') then
      mDerog := Q.FindField('IA_MONTANTECO').AsFloat+Q.FindField('IA_CESSIONECO').AsFloat
          - Q.FindField('IA_MONTANTFISCAL').AsFloat-Q.FindField('IA_CESSIONFISCAL').AsFloat
    else mDerog := 0;
    if (Q.FindField('I_METHODEECO').AsString = Methode) or (Methode = '') then result := result +
        Q.FindField('IA_MONTANTECO').AsFloat+Q.FindField('IA_CESSIONECO').AsFloat
    else if (Methode = 'DDE') and (mDerog < 0) then result := result + (-1)*mDerog
    else if (Methode = 'RDE') and (mDerog > 0) then result := result + mDerog
    else if (Methode='EXC') then result := result + Q.FindField('I_MONTANTEXC').AsFloat+ Q.FindField('I_MONTANTEXCCED').AsFloat;

    // ajout mbo pour exceptionnel lors de la cession
    if (Methode='EXC') and (Q.FindField('I_OPECESSION').AsString = 'X') then
    begin
      sSelect2 :='Select IL_MONTANTEXC, IL_TYPEEXC, I_PLANACTIF FROM IMMOLOG LEFT JOIN IMMO ON (I_IMMO=IL_IMMO)' ;
      sWhere2 :='WHERE IL_PLANACTIFAP = I_PLANACTIF';
      Q2:=OpenSQL (sSelect2+sWhere2,True);
      if (Q2.FindField('IL_TYPEEXC').AsString = 'DOT') then
         result := result + Q2.FindField('IL_MONTANTEXC').AsFloat
      else
         result := result - Q2.FindField('IL_MONTANTEXC').AsFloat;
      Ferme(Q2);
    end;
    Q.Next;
  end ;
  Ferme (Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionAmort(CpteInf,CpteSup : string) : Variant ;
var Q : TQuery;
    sSelect, sWhereCompte : string;
    CodeEnCours : string;
begin
  //!!!! manque reprise aux amortissements.
  AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT I_IMMO,IA_IMMO,IA_CESSIONECO,I_OPECESSION,I_PLANACTIF,I_REPCEDECO FROM IMMO LEFT JOIN IMMOAMOR ON (I_IMMO=IA_IMMO) WHERE I_ETAT<>"FER" AND (I_OPECESSION="X") AND (IA_NUMEROSEQ=I_PLANACTIF OR IA_NUMEROSEQ IS NULL OR IA_NUMEROSEQ="")';
  sWhereCompte := ' AND (I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'")  ORDER BY I_IMMO,IA_DATE ASC';
  Q:=OpenSQL (sSelect+sWhereCompte,True);
  Result:=0 ; CodeEnCours:='' ;
  while not Q.Eof do
  begin
    if CodeEnCours <> Q.FindField('I_IMMO').AsString then result := result+Q.FindField('I_REPCEDECO').AsFloat;
{$IFNDEF EAGLCLIENT}
    if not Q.FindField('IA_CESSIONECO').IsNull then
{$ENDIF}
    Result:=Arrondi(Result+Q.FindField('IA_CESSIONECO').AsFloat,V_PGI.OkDecV) ;
    CodeEnCours:=Q.FindField('I_IMMO').AsString ;
    Q.Next;
  end;
  Ferme (Q);
end ;

Function AmGetListeLieuImmo : Variant ;
var Q : TQuery;
  stLieuGeo : string;
begin
  stLieuGeo := '';
  Q:=OpenSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GEO" ORDER BY CC_CODE' ,True);
  while not Q.Eof do
  begin
    stLieuGeo := stLieuGeo+Q.FindField('CC_CODE').AsString+'='+
      Q.FindField('CC_LIBELLE').AsString+';';
    Q.Next;
  end;
  Ferme (Q);
  Result := stLieuGeo;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/02/2006
Modifié le ... :   /  /    
Description .. : Retourne la somme des valeurs des immobilisations en 
Suite ........ : fonction de leur appartenance à un regroupement.
Mots clefs ... : 
*****************************************************************}
function AmGetValHTImmo ( pstCpteInf, pstCpteSup : string; pstParam : string ) : variant;
var
  lQ : TQuery;
  lstCondition : string;
  lstWhereCompte : string;
begin
  AmBourreBornes (pstCpteInf, pstCpteSup );
  lstWhereCompte := ' (I_COMPTEIMMO>="'+pstCpteInf+'" AND I_COMPTEIMMO <="'+pstCpteSup+'") AND I_OPECESSION<>"X" ';

  // TGA 30/03/2006  if pstParam = 'C' then lstCondition := ' AND I_GROUPEIMMO<>"" '

  if pstParam = 'C' then
      lQ := OpenSQL ('SELECT SUM(I_MONTANTHT) SHT FROM IMMO LEFT JOIN CHOIXCOD ON (I_GROUPEIMMO=CC_CODE) WHERE '+
          lstWhereCompte+ ' AND CC_TYPE = "IGI" AND CC_LIBRE="X" ',True)
  Else
    Begin
      if pstParam = 'N' then
        lstCondition := ' AND (I_GROUPEIMMO="" OR (I_GROUPEIMMO is NULL))'
      else
        lstCondition := '';
      lQ := OpenSQL ('SELECT SUM(I_MONTANTHT) SHT FROM IMMO WHERE '+lstWhereCompte+lstCondition,True);
    end;

  try
    if not lQ.Eof then Result := Arrondi(lQ.FindField('SHT').AsFloat,V_PGI.OkDecV)
    else Result := 0;
  finally
    Ferme (lQ);
  end;
end;

// MVG 12/01/2007
//"DDD"     => Dérogatoire Dotations Différentiel de durée
//"DDM"    => Dérogatoire Dotations Mode dégressif
//"DDA"     => Dérogatoire Dotations Amortissement fiscal exceptionnel
//"DRD"     => Dérogatoire Reprises Différentiel de durée
//"DRM"    => Dérogatoire Reprises Mode dégressif
//"DRA"     => Dérogatoire Reprises Amortissement fiscal exceptionnel
// FQ 19514

Function AmGetDerogatoire(CpteInf,CpteSup, Methode : string) : Variant ;
var Q : TQuery;
    sSelect, sWhere , sWhereCompte, sWhereDerog : string;
    mDerog : double;
    PlanInfo : TPlanInfo;
    DateRef : TDateTime;
    CodeImmo : string;
    cas : integer;
    DotRep : string;
begin
  AmBourreBornes(CpteInf,CpteSup);
  PlanInfo := TPlanInfo.Create('');
  sSelect:='Select I_IMMO,I_DUREEECO,I_DUREEFISC, I_METHODEECO,I_METHODEFISC FROM IMMO ' ;
  // CA - 04/09/2002 - On ne prend en compte que les immos en pleine propriété

  sWhere:='WHERE I_ETAT<>"FER" AND I_NATUREIMMO="PRO" AND I_METHODEFISC<>"" AND I_QUALIFIMMO="R"' ;
  sWhereCompte:=' AND (I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'")' ;

  // FQ 20256 Améliorer les perfs en filtrant le SELECT sur le dérogatoire
  sWhereDerog := '';
  if (Methode = 'DDD') or (Methode = 'DRD') then sWhereDerog := ' AND I_TYPEDEROGLIA = "DUR"';
  if (Methode = 'DDM') or (Methode = 'DRM') then sWhereDerog := ' AND I_TYPEDEROGLIA = "DEG"';
  if (Methode = 'DDA') or (Methode = 'DRA') then sWhereDerog := ' AND I_TYPEDEROGLIA = "EXC"';
  Q:=OpenSQL (sSelect+sWhere+sWhereCompte+sWhereDerog,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    CodeImmo:=Q.FindField('I_IMMO').AsString;
    PlanInfo.ChargeImmo(CodeImmo);
    DateRef:=VH^.Encours.Fin;
    PlanInfo.Calcul(DateRef, true, false, '');  // mbo - 29/10/07 - fq 21754
    DotRep:=copy(methode,2,1);
    if (PlanInfo.GetTypeSortie(DateRef)<>'') and (DotRep='R') then
    begin
             mderog := PlanInfo.GetCumulAntDerogatoire(DateRef);  // fq 17512 chantier fiscal
             if PlanInfo.Derogatoire>0 then mderog:=mderog+PlanInfo.Derogatoire;
//             if PlanInfo.Derogatoire<0 then mderog:=mderog-((-1)*PlanInfo.Derogatoire);
//             if PlanInfo.Derogatoire<0 then mderog:=mderog+((-1)*PlanInfo.Derogatoire);
             mDerog := (-1) * mDerog
    end
    else
    mDerog:=PlanInfo.Derogatoire;

    if PlanInfo.Plan.AmortEco.Duree > PlanInfo.Plan.AmortFisc.Duree then cas:=1
    else
        if PlanInfo.Plan.AmortFisc.Methode='DEG' then cas:=2
        else
        cas:=3;
  if (mDerog>0) and (dotrep='D') then
  begin
     If (Methode='DDD') and (cas=1) then Result:=result + mDerog;
     If (Methode='DDM') and (cas=2) then Result:=result + mDerog;
     If (Methode='DDA') and (cas=3) then Result:=result + mDerog;
  end;
  if (mDerog<0) and (DotRep='R') then
  begin
     If (Methode='DRD') and (cas=1) then Result:=result + (-1)*mDerog;
     If (Methode='DRM') and (cas=2) then Result:=result + (-1)*mDerog;
     If (Methode='DRA') and (cas=3) then Result:=result + (-1)*mDerog;
  end;

   Q.Next;
  end ;
  Ferme (Q);
  Planinfo.free;
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

{
NE SERT PAS - MVG 16/01/2007

Function AmGetCumulDerogatoire(CpteInf,CpteSup : string) : Variant ;
var Q : TQuery;
    sSelect, sWhere , sWhereCompte : string;
    mDerog : double;
    PlanInfo : TPlanInfo;
    DateRef : TDateTime;
    CodeImmo : string;
begin
  AmBourreBornes(CpteInf,CpteSup);
  PlanInfo := TPlanInfo.Create('');
  sSelect:='Select I_IMMO,I_DUREEECO,I_DUREEFISC, I_METHODEECO,I_METHODEFISC FROM IMMO ' ;
  // CA - 04/09/2002 - On ne prend en compte que les immos en pleine propriété

  sWhere:='WHERE I_ETAT<>"FER" AND I_NATUREIMMO="PRO" AND I_METHODEFISC<>"" AND I_QUALIFIMMO="R"' ;
  sWhereCompte:=' AND (I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'")' ;
  Q:=OpenSQL (sSelect+sWhere+sWhereCompte,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    CodeImmo:=Q.FindField('I_IMMO').AsString;
    PlanInfo.ChargeImmo(CodeImmo);
    DateRef:=VH^.Encours.Fin;
    PlanInfo.Calcul(DateRef, true);
    if (PlanInfo.GetTypeSortie(DateRef)<>'') then
       mderog:=0.0
    else
       mderog:=PlanInfo.GetCumulAntDerogatoire + PlanInfo.Derogatoire;
    Result:=result + mDerog;
   Q.Next;
  end ;
  Ferme (Q);
  Planinfo.free;
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ; }

{***********A.G.L.***********************************************
Auteur  ...... : CMX
Créé le ...... : 29/05/2007
Modifié le ... :   /  /
Description .. : Adaptation de la fonction
Suite ........ : THalleyWindow.ChargeLesImmos de HALOle pour Etafi
Mots clefs ... :
*****************************************************************}

function AmGetLesImmos(stType: string; iVersion: integer; laTob: TOB): TOB;
var
  OBImmo, OBAmor, LaTobFille: TOB;
  PlanInfo: TPlanInfo;
  PMV: TImPMValue;
  i: integer;
  strRequete :string;

  function VirePV(St: string):string;
  begin
    Result := FindEtReplace(St, ';', ' ', TRUE);
  end;
begin
  OBImmo := TOB.Create('LES IMMOS', nil, -1);
  if (stType = 'ENC') then
  begin
    PlanInfo := TPlanInfo.Create('');
    strRequete := 'SELECT I_PLANACTIF,I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_DUREEREPRISE,' +
      'I_DATEPIECEA,I_DATEAMORT,I_BASEECO,I_BASEFISC,I_BASEAMORDEBEXO,I_BASEAMORFINEXO,' +
      'I_DATECREATION,I_MONTANTEXC,I_TYPEEXC,I_REINTEGRATION,I_QUOTEPART,I_TVARECUPERABLE, I_TVARECUPEREE, ' +
      'I_VALEURACHAT,I_ETAT,I_MONTANTHT,I_REINTEGRATION,I_OPECHANGEPLAN,I_DATECESSION,' +
      'I_REVISIONECO,I_REPRISEECO,I_REPCEDECO,I_BASEECO,I_DUREEECO,I_METHODEECO,' +
      'I_TAUXECO,I_METHODEFISC,I_REVISIONFISCALE,I_REPRISEFISCAL,I_REPCEDFISC,' +
      'I_BASEFISC,I_DUREEFISC,I_METHODEFISC,I_TAUXFISC,I_OPECESSION FROM IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_ETAT<>"FER" AND I_QUALIFIMMO="R" AND I_DATEPIECEA<="' + USDatetime(VH^.Encours.Fin) + '"'; //, True);
    OBImmo.LoadDetailFromSQL (strRequete);
    for i := 0 to OBImmo.Detail.Count - 1 do
    begin
// Chargement du plan d'amortissement
      strRequete := 'SELECT IA_BASEDEBEXOCESS,IA_BASEDEBEXOECO,IA_BASEDEBEXOFISC,IA_CESSIONARD,IA_CESSIONDEROG,IA_CESSIONDPI,IA_CESSIONECO,IA_CESSIONFISCAL,IA_CESSIONND,IA_CESSIONSBV,' +
                    'IA_CHANGEAMOR,IA_COMPTEIMMO,IA_DATE,IA_DATE1,IA_DATE2,IA_DATE3,IA_DOUBLE1,IA_DOUBLE2,IA_DOUBLE3,IA_IMMO,IA_MONTANTARD,IA_MONTANTDEROG,IA_MONTANTDPI,IA_MONTANTECO,' +
                    'IA_MONTANTFISCAL,IA_MONTANTSBV,IA_NONDEDUCT,IA_NUMEROSEQ,IA_QUOTEPART,IA_REINTEGRATION,IA_STRING1,IA_STRING2,IA_STRING3' +
                    ' FROM IMMOAMOR WHERE IA_IMMO="' + OBImmo.Detail[i].GetValue('I_IMMO') + '" ORDER BY IA_NUMEROSEQ DESC, IA_DATE'; //, True);
      OBAmor := TOB.Create('LES IMMOS', OBImmo.Detail[i], -1);
      OBAmor.LoadDetailFromSQL (strRequete); //LoadDetailDB('IMMOAMOR', '', '', Q, False, False);
// Calcul des informations
      PlanInfo.ChargeImmo (OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul(VH^.Encours.Fin, false, false, '');    //mbo - 29/10/07 - fq 21754
     // DecodeDate(OBImmo.Detail[i].GetValue('I_DATEPIECEA'), a, m, j);
      LaTobFille := TOB.Create('TOB VALEUR', laTob, -1);
      LaTobFille.AddChampSupValeur ('I_INDICE', i + 1); //indice de ligne
      LaTobFille.AddChampSupValeur ('I_LIBELLE', VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE')))); //Libelle
      LaTobFille.AddChampSupValeur ('I_DATEPIECEA', OBImmo.Detail[i].GetValue('I_DATEPIECEA')); //j + m + a); //date d'achat
      LaTobFille.AddChampSupValeur ('I_MONTANTHT', OBImmo.Detail[i].GetValue('I_MONTANTHT')); //Montant HT
      LaTobFille.AddChampSupValeur ('I_TVARECUPERABLE', OBImmo.Detail[i].GetValue('I_TVARECUPERABLE')); //TVA récupérable
      LaTobFille.AddChampSupValeur ('I_TVARECUPEREE', OBImmo.Detail[i].GetValue('I_TVARECUPEREE')); //TVA récupérée
      LaTobFille.AddChampSupValeur ('I_METHODEECO', Copy(OBImmo.Detail[i].GetValue('I_METHODEECO'), 1, 1)); //Méthode d'amortissement
      LaTobFille.AddChampSupValeur ('I_TAUXECO', OBImmo.Detail[i].GetValue('I_TAUXECO')); //le taux d'amortissement économique
      LaTobFille.AddChampSupValeur ('I_CUMULANTECO', PlanInfo.CumulAntEco); //montant des amortissements antérieurs
      LaTobFille.AddChampSupValeur ('I_DOTATIONECO', PlanInfo.DotationEco); //la dotation de l'exercice
    end;
    PlanInfo.Free;
    OBImmo.Free;
  end

  else if (stType = 'SOR') then
  begin
    strRequete := 'SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,' +
      'I_METHODEECO,I_REGLECESSION,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES,IL_VOCEDEE' +
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_QUALIFIMMO="R" AND IL_TYPEOP LIKE "CE%" ' +
      ' AND IL_DATEOP>="' + USDateTime(VH^.Encours.Deb) + '"' +
      ' AND IL_DATEOP<="' + USDatetime(VH^.Encours.Fin) + '"';
    OBImmo.LoadDetailFromSQL (strRequete);
    for i := 0 to OBImmo.Detail.Count - 1 do
    begin
      //DecodeDate(OBImmo.Detail[i].GetValue('I_DATEPIECEA'), a, m, j);
      //DecodeDate(OBImmo.Detail[i].GetValue('IL_DATEOP'), a1, m1, j1);
      PMV := CalculPMValue(OBImmo.Detail[i].GetValue('IL_DATEOP'),
        OBImmo.Detail[i].GetValue('I_DATEPIECEA'), OBImmo.Detail[i].GetValue('I_DATEAMORT'),
        OBImmo.Detail[i].GetValue('IL_PVALUE'), OBImmo.Detail[i].GetValue('IL_CUMANTCESECO'),
        OBImmo.Detail[i].GetValue('IL_DOTCESSECO'), OBImmo.Detail[i].GetValue('I_METHODEECO'),
        OBImmo.Detail[i].GetValue('I_REGLECESSION'));

      LaTobFille := TOB.Create('', laTob, -1);
      LaTobFille.AddChampSupValeur('I_INDICE', i + 1); //indice de ligne
      LaTobFille.AddChampSupValeur('I_LIBELLE', VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE')))); //Libelle
      LaTobFille.AddChampSupValeur('I_DATEOP', OBImmo.Detail[i].GetValue('IL_DATEOP'));//j1 + m1 + a1); //date de sortie ou date d'opération
      LaTobFille.AddChampSupValeur('I_DATEPIECEA', OBImmo.Detail[i].GetValue('I_DATEPIECEA')); //j + m + a); //date d'achat
      LaTobFille.AddChampSupValeur('I_VOCEDEE', OBImmo.Detail[i].GetValue('IL_VOCEDEE')); //montant HT
      LaTobFille.AddChampSupValeur('I_CUMULDOT', OBImmo.Detail[i].GetValue('IL_CUMANTCESECO') + OBImmo.Detail[i].GetValue('IL_DOTCESSECO')); //cumul des amortissements
      LaTobFille.AddChampSupValeur('I_MONTANTCES', OBImmo.Detail[i].GetValue('IL_MONTANTCES')); //prix de sortie
      LaTobFille.AddChampSupValeur('I_PMVCT', PMV.PCT - PMV.MCT); //+/- value à court terme
      LaTobFille.AddChampSupValeur('I_PMVLT', PMV.PLT - PMV.MLT); //+/- valeur à long terme
    end;
  end;
  Result := laTob;
end;


//================== fonctions avec un select sur code etablissement =================================

Function AmGetAugAcqMutEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    sSelect, sWhere : string;
    CodeEnCours : string;
begin
  if Where = '' then //ajout CMX 01/10/2007
        AmBourreBornes (CpteInf,CpteSup);
  CodeEnCours:='' ;
  sSelect:='SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,I_MONTANTHT,I_VALEURACHAT,I_COMPTEREF FROM IMMOLOG LEFT JOIN IMMO ON (I_IMMO=IL_IMMO)';
  sWhere:=' WHERE I_ETAT<>"FER" AND (IL_TYPEOP="ACQ" OR IL_TYPEOP="MUT") AND (IL_DATEOP>="'+
        USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
        +'")';
  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sWhere := sWhere + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhere := sWhere + Where + 'ORDER BY IL_IMMO,IL_DATEOP DESC'
  else
    sWhere := sWhere + ' AND (I_COMPTEREF>="' + CpteInf + '" AND I_COMPTEREF <="' + CpteSup + '") ORDER BY IL_IMMO,IL_DATEOP DESC';

  Q:=OpenSQL(sSelect+sWhere,True) ;
  Result:=0 ;
  while not Q.Eof do
  begin
    if Q.FindField('IL_IMMO').AsString <> CodeEnCours then
    begin
//    result :=  result + Q.FindField('I_MONTANTHT').AsFloat;
      if Q.FindField('IL_TYPEOP').AsString = 'ACQ' then
        result :=  result + Q.FindField('I_VALEURACHAT').AsFloat
      else result := result + Q.FindField('I_MONTANTHT').AsFloat;
      CodeEnCours := Q.FindField('IL_IMMO').AsString;
    end;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionSortieEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    sSelect, sWhere : string;
begin
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT IL_VOCEDEE,IL_TYPEOP,IL_DATEOP,I_COMPTEREF FROM IMMOLOG LEFT JOIN IMMO ON (IL_IMMO=I_IMMO)';
  sWhere := ' WHERE I_ETAT<>"FER"';

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sWhere := sWhere + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  { modif CMX 01/10/2007 sWhere := sWhere + ' AND IL_TYPEOP LIKE "CE%" AND (IL_DATEOP>="'+
        USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
         +'") AND (I_COMPTEREF>="'+CpteInf+'" AND I_COMPTEREF <="'+CpteSup+'")'; }

  sWhere := sWhere + ' AND IL_TYPEOP LIKE "CE%" AND (IL_DATEOP>="' +
            USDateTime(VH^.Encours.Deb) + '" AND IL_DATEOP<="' + USDateTime(VH^.Encours.Fin) + '")';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhere := sWhere + Where
  else
    sWhere := sWhere + ' AND (I_COMPTEREF>="' + CpteInf + '" AND I_COMPTEREF <="' + CpteSup + '")';

  Result:=0 ;
  Q:=OpenSQL (sSelect+sWhere,True);
  while not Q.Eof do
  begin
    Result:=Result+Q.FindField('IL_VOCEDEE').AsFloat ;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionMutationEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    sSelect, sWhere, CodeEnCours : string;
begin
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,IL_CPTEMUTATION,I_COMPTEREF,I_MONTANTHT FROM IMMOLOG LEFT JOIN IMMO ON (IL_IMMO=I_IMMO)';
  sWhere := ' WHERE I_ETAT<>"FER"';

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sWhere := sWhere + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  { modif CMX 01/10/2007 sWhere := sWhere + ' AND IL_TYPEOP="MUT" AND (IL_DATEOP>="'+
          USDateTime(VH^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VH^.Encours.Fin)
          +'") AND (IL_CPTEMUTATION>="'+CpteInf+'" AND IL_CPTEMUTATION <="'+CpteSup+'")'
          +'  ORDER BY IL_IMMO,IL_DATEOP DESC';  }

   sWhere := sWhere + ' AND IL_TYPEOP="MUT" AND (IL_DATEOP>="'
    + USDateTime(VH^.Encours.Deb) + '" AND IL_DATEOP<="' + USDateTime(VH^.Encours.Fin) + '")';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhere := sWhere + Where
  else
    sWhere := sWhere + ' AND (IL_CPTEMUTATION>="' + CpteInf + '" AND IL_CPTEMUTATION <="' + CpteSup + '")';

  sWhere := sWhere + '  ORDER BY IL_IMMO,IL_DATEOP DESC';

// MVG FQ 18123 +' AND (I_COMPTEREF<"'+CpteInf+'" OR I_COMPTEREF >"'+CpteSup+'") ORDER BY IL_IMMO,IL_DATEOP DESC';
  Q:=OpenSQL(sSelect+sWhere,True) ;
  Result:=0 ; CodeEnCours:='' ;
  while not Q.Eof do
  begin
    if CodeEnCours <> Q.FindField('IL_IMMO').AsString then
    begin
      Result:=Result+Q.FindField('I_MONTANTHT').AsFloat ;
      CodeEnCours:=Q.FindField('IL_IMMO').AsString ;
    end;
    Q.Next;
  end;
  Ferme(Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDotationExoEtab(CpteInf,CpteSup, Methode, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    Q2: TQuery;
    sSelect, sWhere , sWhereCompte : string;
    sSelect2, sWhere2 : string;
    mDerog : double;
begin
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes(CpteInf,CpteSup);
  sSelect:='Select IA_MONTANTECO,IA_MONTANTFISCAL,IA_CESSIONECO,IA_CESSIONFISCAL,I_PLANACTIF,I_METHODEECO,I_METHODEFISC,I_MONTANTEXC,I_MONTANTEXCCED, I_OPECESSION FROM IMMOAMOR LEFT JOIN IMMO ON (I_IMMO=IA_IMMO)' ;
  // CA - 04/09/2002 - On ne prend en compte que les immos en pleine propriÚtÚ
  sWhere:='WHERE I_ETAT<>"FER" AND I_NATUREIMMO="PRO"';

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sWhere := sWhere + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  { modif CMX 01/10/2007 sWhere := sWhere + ' AND (IA_NUMEROSEQ=I_PLANACTIF AND IA_DATE="'+UsDateTime(VH^.Encours.Fin)+'")' ;
  sWhereCompte:=' AND (IA_COMPTEIMMO>="'+CpteInf+'" AND IA_COMPTEIMMO <="'+CpteSup+'")' ;  }

  sWhere := sWhere + ' AND (IA_NUMEROSEQ=I_PLANACTIF AND IA_DATE="' + UsDateTime(VH^.Encours.Fin) + '")';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhereCompte := Where
  else
    sWhereCompte := ' AND (IA_COMPTEIMMO>="' + CpteInf + '" AND IA_COMPTEIMMO <="' + CpteSup + '")';

  Q:=OpenSQL (sSelect+sWhere+sWhereCompte,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    if (Q.FindField('I_METHODEFISC').AsString <> '') then
      mDerog := Q.FindField('IA_MONTANTECO').AsFloat+Q.FindField('IA_CESSIONECO').AsFloat
          - Q.FindField('IA_MONTANTFISCAL').AsFloat-Q.FindField('IA_CESSIONFISCAL').AsFloat
    else mDerog := 0;
    if (Q.FindField('I_METHODEECO').AsString = Methode) or (Methode = '') then result := result +
        Q.FindField('IA_MONTANTECO').AsFloat+Q.FindField('IA_CESSIONECO').AsFloat
    else if (Methode = 'DDE') and (mDerog < 0) then result := result + (-1)*mDerog
    else if (Methode = 'RDE') and (mDerog > 0) then result := result + mDerog
    else if (Methode='EXC') then result := result + Q.FindField('I_MONTANTEXC').AsFloat+ Q.FindField('I_MONTANTEXCCED').AsFloat;

    // ajout mbo pour exceptionnel lors de la cession
    if (Methode='EXC') and (Q.FindField('I_OPECESSION').AsString = 'X') then
    begin
      sSelect2 :='Select IL_MONTANTEXC, IL_TYPEEXC, I_PLANACTIF FROM IMMOLOG LEFT JOIN IMMO ON (I_IMMO=IL_IMMO)' ;
      sWhere2 :='WHERE IL_PLANACTIFAP = I_PLANACTIF';
      Q2:=OpenSQL (sSelect2+sWhere2,True);
      if (Q2.FindField('IL_TYPEEXC').AsString = 'DOT') then
         result := result + Q2.FindField('IL_MONTANTEXC').AsFloat
      else
         result := result - Q2.FindField('IL_MONTANTEXC').AsFloat;
      Ferme(Q2);
    end;
    Q.Next;
  end ;
  Ferme (Q);
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;

Function AmGetDiminutionAmortEtab(CpteInf,CpteSup, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    sSelect, sWhereCompte : string;
    CodeEnCours : string;
begin
  //!!!! manque reprise aux amortissements.
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes (CpteInf,CpteSup);
  sSelect := 'SELECT I_IMMO,I_ETABLISSEMENT, IA_IMMO,IA_CESSIONECO,I_OPECESSION,I_PLANACTIF,I_REPCEDECO FROM IMMO LEFT JOIN IMMOAMOR ON (I_IMMO=IA_IMMO) WHERE I_ETAT<>"FER"';

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sSelect := sSelect + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  sSelect := sSelect + ' AND (I_OPECESSION="X") AND (IA_NUMEROSEQ=I_PLANACTIF OR IA_NUMEROSEQ IS NULL OR IA_NUMEROSEQ="")';

  // modif CMX 01/10/2007 sWhereCompte := ' AND (I_COMPTEIMMO>="'+CpteInf+'" AND I_COMPTEIMMO <="'+CpteSup+'")  ORDER BY I_IMMO,IA_DATE ASC';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhereCompte := Where + ' ORDER BY I_IMMO,IA_DATE ASC'
  else
    sWhereCompte := ' AND (I_COMPTEIMMO>="' + CpteInf + '" AND I_COMPTEIMMO <="' + CpteSup + '") ORDER BY I_IMMO,IA_DATE ASC';


  Q:=OpenSQL (sSelect+sWhereCompte,True);
  Result:=0 ; CodeEnCours:='' ;
  while not Q.Eof do
  begin
    if CodeEnCours <> Q.FindField('I_IMMO').AsString then result := result+Q.FindField('I_REPCEDECO').AsFloat;
{$IFNDEF EAGLCLIENT}
    if not Q.FindField('IA_CESSIONECO').IsNull then
{$ENDIF}
    Result:=Arrondi(Result+Q.FindField('IA_CESSIONECO').AsFloat,V_PGI.OkDecV) ;
    CodeEnCours:=Q.FindField('I_IMMO').AsString ;
    Q.Next;
  end;
  Ferme (Q);
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
CrÚÚ le ...... : 06/02/2006
ModifiÚ le ... :   /  /
Description .. : Retourne la somme des valeurs des immobilisations en
Suite ........ : fonction de leur appartenance Ó un regroupement.
Mots clefs ... :
*****************************************************************}
function AmGetValHTImmoEtab ( pstCpteInf, pstCpteSup : string; pstParam, Etablissement : string; Where: string = '') : variant;
var
  lQ : TQuery;
  lstCondition : string;
  lstWhereCompte : string;
begin
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes (pstCpteInf, pstCpteSup );

  // CMX 01/10/2007 lstWhereCompte := ' (I_COMPTEIMMO>="'+pstCpteInf+'" AND I_COMPTEIMMO <="'+pstCpteSup+'") AND I_OPECESSION<>"X" ';
  if Where <> '' then //ajout CMX 01/10/2007
    lstWhereCompte := Where + ' AND I_OPECESSION<>"X" '
  else
    lstWhereCompte := ' (I_COMPTEIMMO>="' + pstCpteInf + '" AND I_COMPTEIMMO <="' + pstCpteSup + '") AND I_OPECESSION<>"X" ';

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     lstWhereCompte := lstWhereCompte + 'AND (I_ETABLISSEMENT ="' + Etablissement +'") ';


  // TGA 30/03/2006  if pstParam = 'C' then lstCondition := ' AND I_GROUPEIMMO<>"" '

  if pstParam = 'C' then
      lQ := OpenSQL ('SELECT SUM(I_MONTANTHT) SHT FROM IMMO LEFT JOIN CHOIXCOD ON (I_GROUPEIMMO=CC_CODE) WHERE '+
          lstWhereCompte+ ' AND CC_TYPE = "IGI" AND CC_LIBRE="X" ',True)
  Else
    Begin
      if pstParam = 'N' then
        lstCondition := ' AND (I_GROUPEIMMO="" OR (I_GROUPEIMMO is NULL))'
      else
        lstCondition := '';
      lQ := OpenSQL ('SELECT SUM(I_MONTANTHT) SHT FROM IMMO WHERE '+lstWhereCompte+lstCondition,True);
    end;

  try
    if not lQ.Eof then Result := Arrondi(lQ.FindField('SHT').AsFloat,V_PGI.OkDecV)
    else Result := 0;
  finally
    Ferme (lQ);
  end;
end;

// MVG 12/01/2007
//"DDD"     => DÚrogatoire Dotations DiffÚrentiel de durÚe
//"DDM"    => DÚrogatoire Dotations Mode dÚgressif
//"DDA"     => DÚrogatoire Dotations Amortissement fiscal exceptionnel
//"DRD"     => DÚrogatoire Reprises DiffÚrentiel de durÚe
//"DRM"    => DÚrogatoire Reprises Mode dÚgressif
//"DRA"     => DÚrogatoire Reprises Amortissement fiscal exceptionnel
// FQ 19514

Function AmGetDerogatoireEtab(CpteInf,CpteSup, Methode, Etablissement : string; Where: string = '') : Variant ;
var Q : TQuery;
    sSelect, sWhere , sWhereCompte, sWhereDerog : string;
    mDerog : double;
    PlanInfo : TPlanInfo;
    DateRef : TDateTime;
    CodeImmo : string;
    cas : integer;
    DotRep : string;
begin
  if Where = '' then //ajout CMX 01/10/2007
     AmBourreBornes(CpteInf,CpteSup);
  PlanInfo := TPlanInfo.Create('');
  sSelect:='Select I_IMMO,I_DUREEECO,I_DUREEFISC, I_METHODEECO,I_METHODEFISC, I_ETABLISSEMENT FROM IMMO ' ;
  // CA - 04/09/2002 - On ne prend en compte que les immos en pleine propriÚtÚ

  sWhere:='WHERE I_ETAT<>"FER" AND I_NATUREIMMO="PRO" AND I_METHODEFISC<>"" AND I_QUALIFIMMO="R"' ;

  // ajout mbo 20.07.07
  if Etablissement <> '' then
     sWhere := sWhere + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

  if Where <> '' then //ajout CMX 01/10/2007
    sWhereCompte := Where
  else
    sWhereCompte := ' AND (I_COMPTEIMMO>="' + CpteInf + '" AND I_COMPTEIMMO <="' + CpteSup + '")';


  // FQ 20256 AmÚliorer les perfs en filtrant le SELECT sur le dÚrogatoire
  sWhereDerog := '';
  if (Methode = 'DDD') or (Methode = 'DRD') then sWhereDerog := ' AND I_TYPEDEROGLIA = "DUR"';
  if (Methode = 'DDM') or (Methode = 'DRM') then sWhereDerog := ' AND I_TYPEDEROGLIA = "DEG"';
  if (Methode = 'DDA') or (Methode = 'DRA') then sWhereDerog := ' AND I_TYPEDEROGLIA = "EXC"';
  Q:=OpenSQL (sSelect+sWhere+sWhereCompte+sWhereDerog,True);
  Result:=0 ;
  while not Q.Eof do
  begin
    CodeImmo:=Q.FindField('I_IMMO').AsString;
    PlanInfo.ChargeImmo(CodeImmo);
    DateRef:=VH^.Encours.Fin;
    PlanInfo.Calcul(DateRef, true, false, '');    //mbo - 29/10/07 - fq 21754
    DotRep:=copy(methode,2,1);
    if (PlanInfo.GetTypeSortie(DateRef)<>'') and (DotRep='R') then
    begin
             mderog := PlanInfo.GetCumulAntDerogatoire(DateRef);  // fq 17512 chantier fiscal
             if PlanInfo.Derogatoire>0 then mderog:=mderog+PlanInfo.Derogatoire;
//             if PlanInfo.Derogatoire<0 then mderog:=mderog-((-1)*PlanInfo.Derogatoire);
//             if PlanInfo.Derogatoire<0 then mderog:=mderog+((-1)*PlanInfo.Derogatoire);
             mDerog := (-1) * mDerog
    end
    else
    mDerog:=PlanInfo.Derogatoire;

    if PlanInfo.Plan.AmortEco.Duree > PlanInfo.Plan.AmortFisc.Duree then cas:=1
    else
        if PlanInfo.Plan.AmortFisc.Methode='DEG' then cas:=2
        else
        cas:=3;
  if (mDerog>0) and (dotrep='D') then
  begin
     If (Methode='DDD') and (cas=1) then Result:=result + mDerog;
     If (Methode='DDM') and (cas=2) then Result:=result + mDerog;
     If (Methode='DDA') and (cas=3) then Result:=result + mDerog;
  end;
  if (mDerog<0) and (DotRep='R') then
  begin
     If (Methode='DRD') and (cas=1) then Result:=result + (-1)*mDerog;
     If (Methode='DRM') and (cas=2) then Result:=result + (-1)*mDerog;
     If (Methode='DRA') and (cas=3) then Result:=result + (-1)*mDerog;
  end;

   Q.Next;
  end ;
  Ferme (Q);
  Planinfo.free;
  Result:=Arrondi(Result,V_PGI.OkDecV) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : CMX
CrÚÚ le ...... : 29/05/2007
ModifiÚ le ... :   /  /
Description .. : Adaptation de la fonction
Suite ........ : THalleyWindow.ChargeLesImmos de HALOle pour Etafi
Mots clefs ... :
*****************************************************************}

function AmGetLesImmosEtab(stType: string; iVersion: integer; laTob: TOB; Etablissement: string; Where: string = ''): TOB;
var
  OBImmo, OBAmor, LaTobFille: TOB;
  PlanInfo: TPlanInfo;
  PMV: TImPMValue;
  i: integer;
  strRequete :string;

  function VirePV(St: string):string;
  begin
    Result := FindEtReplace(St, ';', ' ', TRUE);
  end;
begin
  OBImmo := TOB.Create('LES IMMOS', nil, -1);
  if (stType = 'ENC') then
  begin
    PlanInfo := TPlanInfo.Create('');
    strRequete := 'SELECT I_PLANACTIF,I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_DUREEREPRISE,' +
      'I_DATEPIECEA,I_DATEAMORT,I_BASEECO,I_BASEFISC,I_BASEAMORDEBEXO,I_BASEAMORFINEXO,' +
      'I_DATECREATION,I_MONTANTEXC,I_TYPEEXC,I_REINTEGRATION,I_QUOTEPART,I_TVARECUPERABLE, I_TVARECUPEREE, ' +
      'I_VALEURACHAT,I_ETAT,I_MONTANTHT,I_REINTEGRATION,I_OPECHANGEPLAN,I_DATECESSION,' +
      'I_REVISIONECO,I_REPRISEECO,I_REPCEDECO,I_BASEECO,I_DUREEECO,I_METHODEECO,' +
      'I_TAUXECO,I_METHODEFISC,I_REVISIONFISCALE,I_REPRISEFISCAL,I_REPCEDFISC,' +
      'I_BASEFISC,I_DUREEFISC,I_METHODEFISC,I_TAUXFISC,I_OPECESSION, I_ETABLISSEMENT FROM IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_ETAT<>"FER" AND I_QUALIFIMMO="R" AND I_DATEPIECEA<="' + USDatetime(VH^.Encours.Fin) + '"'; //, True);

    // ajout mbo 20.07.07
    if Etablissement <> '' then
       strRequete := strRequete + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

    if Where <> '' then //ajout CMX 01/10/2007
      strRequete := strRequete + Where;

    OBImmo.LoadDetailFromSQL (strRequete);
    for i := 0 to OBImmo.Detail.Count - 1 do
    begin
// Chargement du plan d'amortissement
      strRequete := 'SELECT IA_BASEDEBEXOCESS,IA_BASEDEBEXOECO,IA_BASEDEBEXOFISC,IA_CESSIONARD,IA_CESSIONDEROG,IA_CESSIONDPI,IA_CESSIONECO,IA_CESSIONFISCAL,IA_CESSIONND,IA_CESSIONSBV,' +
                    'IA_CHANGEAMOR,IA_COMPTEIMMO,IA_DATE,IA_DATE1,IA_DATE2,IA_DATE3,IA_DOUBLE1,IA_DOUBLE2,IA_DOUBLE3,IA_IMMO,IA_MONTANTARD,IA_MONTANTDEROG,IA_MONTANTDPI,IA_MONTANTECO,' +
                    'IA_MONTANTFISCAL,IA_MONTANTSBV,IA_NONDEDUCT,IA_NUMEROSEQ,IA_QUOTEPART,IA_REINTEGRATION,IA_STRING1,IA_STRING2,IA_STRING3' +
                    ' FROM IMMOAMOR WHERE IA_IMMO="' + OBImmo.Detail[i].GetValue('I_IMMO') + '" ORDER BY IA_NUMEROSEQ DESC, IA_DATE'; //, True);
      OBAmor := TOB.Create('LES IMMOS', OBImmo.Detail[i], -1);
      OBAmor.LoadDetailFromSQL (strRequete); //LoadDetailDB('IMMOAMOR', '', '', Q, False, False);
// Calcul des informations
      PlanInfo.ChargeImmo (OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul(VH^.Encours.Fin, false, false, ''); //mbo - 29/10/07 - fq 21754
     // DecodeDate(OBImmo.Detail[i].GetValue('I_DATEPIECEA'), a, m, j);
      LaTobFille := TOB.Create('TOB VALEUR', laTob, -1);
      LaTobFille.AddChampSupValeur ('I_INDICE', i + 1); //indice de ligne
      LaTobFille.AddChampSupValeur ('I_LIBELLE', VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE')))); //Libelle
      LaTobFille.AddChampSupValeur ('I_DATEPIECEA', OBImmo.Detail[i].GetValue('I_DATEPIECEA')); //j + m + a); //date d'achat
      LaTobFille.AddChampSupValeur ('I_MONTANTHT', OBImmo.Detail[i].GetValue('I_MONTANTHT')); //Montant HT
      LaTobFille.AddChampSupValeur ('I_TVARECUPERABLE', OBImmo.Detail[i].GetValue('I_TVARECUPERABLE')); //TVA rÚcupÚrable
      LaTobFille.AddChampSupValeur ('I_TVARECUPEREE', OBImmo.Detail[i].GetValue('I_TVARECUPEREE')); //TVA rÚcupÚrÚe
      LaTobFille.AddChampSupValeur ('I_METHODEECO', Copy(OBImmo.Detail[i].GetValue('I_METHODEECO'), 1, 1)); //MÚthode d'amortissement
      LaTobFille.AddChampSupValeur ('I_TAUXECO', OBImmo.Detail[i].GetValue('I_TAUXECO')); //le taux d'amortissement Úconomique
      LaTobFille.AddChampSupValeur ('I_CUMULANTECO', PlanInfo.CumulAntEco); //montant des amortissements antÚrieurs
      LaTobFille.AddChampSupValeur ('I_DOTATIONECO', PlanInfo.DotationEco); //la dotation de l'exercice
    end;
    PlanInfo.Free;
    OBImmo.Free;
  end

  else if (stType = 'SOR') then
  begin
    strRequete := 'SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,' +
      'I_METHODEECO,I_REGLECESSION,I_ETABLISSEMENT, IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES,IL_VOCEDEE' +
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_QUALIFIMMO="R"';

    // ajout mbo 20.07.07
    if Etablissement <> '' then
       strRequete := strRequete + ' AND (I_ETABLISSEMENT ="' + Etablissement +'")';

    if Where <> '' then //ajout CMX 01/10/2007
      strRequete := strRequete + Where;

    strRequete := strRequete + ' AND IL_TYPEOP LIKE "CE%" ' +
      ' AND IL_DATEOP>="' + USDateTime(VH^.Encours.Deb) + '"' +
      ' AND IL_DATEOP<="' + USDatetime(VH^.Encours.Fin) + '"';
    OBImmo.LoadDetailFromSQL (strRequete);
    for i := 0 to OBImmo.Detail.Count - 1 do
    begin
      //DecodeDate(OBImmo.Detail[i].GetValue('I_DATEPIECEA'), a, m, j);
      //DecodeDate(OBImmo.Detail[i].GetValue('IL_DATEOP'), a1, m1, j1);
      PMV := CalculPMValue(OBImmo.Detail[i].GetValue('IL_DATEOP'),
        OBImmo.Detail[i].GetValue('I_DATEPIECEA'), OBImmo.Detail[i].GetValue('I_DATEAMORT'),
        OBImmo.Detail[i].GetValue('IL_PVALUE'), OBImmo.Detail[i].GetValue('IL_CUMANTCESECO'),
        OBImmo.Detail[i].GetValue('IL_DOTCESSECO'), OBImmo.Detail[i].GetValue('I_METHODEECO'),
        OBImmo.Detail[i].GetValue('I_REGLECESSION'));

      LaTobFille := TOB.Create('', laTob, -1);
      LaTobFille.AddChampSupValeur('I_INDICE', i + 1); //indice de ligne
      LaTobFille.AddChampSupValeur('I_LIBELLE', VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE')))); //Libelle
      LaTobFille.AddChampSupValeur('I_DATEOP', OBImmo.Detail[i].GetValue('IL_DATEOP'));//j1 + m1 + a1); //date de sortie ou date d'opÚration
      LaTobFille.AddChampSupValeur('I_DATEPIECEA', OBImmo.Detail[i].GetValue('I_DATEPIECEA')); //j + m + a); //date d'achat
      LaTobFille.AddChampSupValeur('I_VOCEDEE', OBImmo.Detail[i].GetValue('IL_VOCEDEE')); //montant HT
      LaTobFille.AddChampSupValeur('I_CUMULDOT', OBImmo.Detail[i].GetValue('IL_CUMANTCESECO') + OBImmo.Detail[i].GetValue('IL_DOTCESSECO')); //cumul des amortissements
      LaTobFille.AddChampSupValeur('I_MONTANTCES', OBImmo.Detail[i].GetValue('IL_MONTANTCES')); //prix de sortie
      LaTobFille.AddChampSupValeur('I_PMVCT', PMV.PCT - PMV.MCT); //+/- value Ó court terme
      LaTobFille.AddChampSupValeur('I_PMVLT', PMV.PLT - PMV.MLT); //+/- valeur Ó long terme
    end;
  // ajout mbo - 20.09.07
    OBImmo.Free;
  end

  else if (stType = '') then
  begin
    PlanInfo := TPlanInfo.Create('');
    strRequete := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_OPECESSION,IIF(I_OPECESSION="x",IL_BASEECOAVMB,I_BASEECO) AS BASE,I_TAUXECO,' +
      'IL_DATEOP,IL_MONTANTCES,IL_CUMANTCESECO,IL_DOTCESSECO FROM IMMO LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO AND I_OPECESSION = "X"';
    if Etablissement <> '' then
      strRequete := strRequete + ' AND I_ETABLISSEMENT ="' + Etablissement + '"';

    if Where <> '' then //ajout CMX 01/10/2007
      strRequete := strRequete + Where;

    strRequete := strRequete + ' AND IL_TYPEOP LIKE "CE%" ' +
      ' AND IL_DATEOP>="' + USDateTime(VH^.Encours.Deb) + '"' +
      ' AND IL_DATEOP<="' + USDatetime(VH^.Encours.Fin) + '")' +
      ' ORDER BY I_LIBELLE';
    OBImmo.LoadDetailFromSQL(strRequete);
    for i := 0 to OBImmo.Detail.Count - 1 do
    begin
      //Chargement du plan d'amortissement
      strRequete := 'SELECT IA_BASEDEBEXOCESS,IA_BASEDEBEXOECO,IA_BASEDEBEXOFISC,IA_CESSIONARD,IA_CESSIONDEROG,IA_CESSIONDPI,IA_CESSIONECO,IA_CESSIONFISCAL,IA_CESSIONND,IA_CESSIONSBV,' +
        'IA_CHANGEAMOR,IA_COMPTEIMMO,IA_DATE,IA_DATE1,IA_DATE2,IA_DATE3,IA_DOUBLE1,IA_DOUBLE2,IA_DOUBLE3,IA_IMMO,IA_MONTANTARD,IA_MONTANTDEROG,IA_MONTANTDPI,IA_MONTANTECO,' +
        'IA_MONTANTFISCAL,IA_MONTANTSBV,IA_NONDEDUCT,IA_NUMEROSEQ,IA_QUOTEPART,IA_REINTEGRATION,IA_STRING1,IA_STRING2,IA_STRING3' +
        ' FROM IMMOAMOR WHERE IA_IMMO="' + OBImmo.Detail[i].GetValue('I_IMMO') + '" ORDER BY IA_NUMEROSEQ DESC, IA_DATE'; //, True);
      OBAmor := TOB.Create('LES IMMOS', OBImmo.Detail[i], -1);
      OBAmor.LoadDetailFromSQL(strRequete);
      //Calcul des informations
      PlanInfo.ChargeImmo(OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul(VH^.Encours.Fin, false, false, '');  //mbo - 29/10/07 - fq 21754

      LaTobFille := TOB.Create('', laTob, -1);
      LaTobFille.AddChampSupValeur('I_INDICE', i + 1); //indice de ligne
      LaTobFille.AddChampSupValeur('I_LIBELLE', VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE')))); //Libelle
      LaTobFille.AddChampSupValeur('I_DATEPIECEA', OBImmo.Detail[i].GetValue('I_DATEPIECEA')); //j + m + a); //date d'achat
      LaTobFille.AddChampSupValeur('I_BASE', OBImmo.Detail[i].GetValue('BASE'));
      LaTobFille.AddChampSupValeur('I_TAUXECO', OBImmo.Detail[i].GetValue('I_TAUXECO')); //le taux d'amortissement Úconomique
      LaTobFille.AddChampSupValeur('I_DATEOP', OBImmo.Detail[i].GetValue('IL_DATEOP'));
      LaTobFille.AddChampSupValeur('I_MONTANTCES', OBImmo.Detail[i].GetValue('IL_MONTANTCES')); //prix de sortie
      LaTobFille.AddChampSupValeur('I_CUMULANTECO', PlanInfo.CumulAntEco); //montant des amortissements ant+rieurs
      if OBImmo.Detail[i].GetValue('I_OPECESSION') = 'X' then
        LaTobFille.AddChampSupValeur('I_DOTATIONECO', OBImmo.Detail[i].GetValue('IL_CUMANTCESECO') + OBImmo.Detail[i].GetValue('IL_DOTCESSECO')) //la dotation de l'exercice
      else
        LaTobFille.AddChampSupValeur('I_DOTATIONECO', PlanInfo.DotationEco); //la dotation de l'exercice
    end;
    PlanInfo.Free;
    OBImmo.Free;
  // fin ajout 20.09.07
  end;
  Result := laTob;
end;

end.

