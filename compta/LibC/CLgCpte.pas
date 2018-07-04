{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 28/06/2004
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit CLgCpte;
interface

uses
  Windows, Messages, SysUtils, Classes,
  MajTable,
  Hctrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Ent1,
  HEnt1,
  uTOB,uEntCommun;

type
  TInfoChgLgCpte = procedure ( Msg : string ) of object;

function FormatCpte (Chaine : String; Bourre : Char; Lg : integer) : String;
procedure ChangeLgCpteGen ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte);
procedure ChangeLgCpteAux ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte );
procedure ChangeLgCpteAna ( Lg : integer; Bourre : Char; pszAxe : String; PO : TInfoChgLgCpte );
function PasDeDoublonChampCompte ( Table , Champ : string; Bourre : Char; Lg : integer; LCpt : TStringList; pszAxe : string = '' ) : boolean;
procedure UpdateChampCompteSpecif ( Table : string; Bourre : Char ; Lg : integer; pszType : String; PO : TInfoChgLgCpte);
procedure UpdateChampCompte ( Table : string ; AChamp : array of string; Bourre : Char; Lg : integer; pszType : String; PO : TInfoChgLgCpte; pszWhere : String ='');
function ParametrageCompteIdentique (T1, T2 : TOB; Table : string ) : boolean;

type
 TYPECOMPTE = record
        Champ : string;
        compte : string;
  end;
implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc;



{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/04/2002
Modifié le ... :   /  /
Description .. : Compare les paramétrages de 2 comptes
Mots clefs ... :
*****************************************************************}
function ParametrageCompteIdentique (T1, T2 : TOB; Table : string ) : boolean;
var bOk : boolean;
begin
  if Table = 'GENERAUX' then
  begin
    bOk := (T1.GetValue('G_NATUREGENE')= T2.GetValue('G_NATUREGENE'))
        and (T1.GetValue('G_COLLECTIF')= T2.GetValue('G_COLLECTIF'))
        and (T1.GetValue('G_FERME')= T2.GetValue('G_FERME'))
        and (T1.GetValue('G_LETTRABLE')= T2.GetValue('G_LETTRABLE'))
        and (T1.GetValue('G_POINTABLE')= T2.GetValue('G_POINTABLE'))
        and (T1.GetValue('G_VENTILABLE1')= T2.GetValue('G_VENTILABLE1'))
        and (T1.GetValue('G_VENTILABLE2')= T2.GetValue('G_VENTILABLE2'))
        and (T1.GetValue('G_VENTILABLE3')= T2.GetValue('G_VENTILABLE3'))
        and (T1.GetValue('G_VENTILABLE4')= T2.GetValue('G_VENTILABLE4'))
        and (T1.GetValue('G_VENTILABLE5')= T2.GetValue('G_VENTILABLE5'));
  end else  if Table = 'TIERS' then
  begin
    bOk := (T1.GetValue('T_NATUREAUXI')= T2.GetValue('T_NATUREAUXI'))
        and (T1.GetValue('T_DEVISE')= T2.GetValue('T_DEVISE'))
        and (T1.GetValue('T_MULTIDEVISE')= T2.GetValue('T_MULTIDEVISE'))
        and (T1.GetValue('T_LETTRABLE')= T2.GetValue('T_LETTRABLE'))
//        and (T1.GetValue('T_TIERS')= T2.GetValue('T_TIERS'))  // CA - 29/11/2002
        and (T1.GetValue('T_FERME')= T2.GetValue('T_FERME'));
  end else begin
  	bOk := (T1.GetValue('S_SENS')= T2.GetValue('S_SENS'))
	      and (T1.GetValue('S_CORRESP1')= T2.GetValue('S_CORRESP1'))
        and (T1.GetValue('S_CORRESP2')= T2.GetValue('S_CORRESP2'))
        and (T1.GetValue('S_FERME')= T2.GetValue('S_FERME'))
        and (T1.GetValue('S_CONFIDENTIEL')= T2.GetValue('S_CONFIDENTIEL'))
        and (T1.GetValue('S_AXE')= T2.GetValue('S_AXE'));
  end;
  Result := bOk;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/04/2002
Modifié le ... : 02/12/2002
Description .. : Contrôle la possibilité de doublon suite à une réduction de la
Suite ........ : longueur des comptes;
Suite ........ : Fonctionne pour les tables TIERS et GENERAUX
Suite ........ : - Modification pour la prise en compte des sections ( axe en 
Suite ........ : paramètre ).
Mots clefs ... : 
*****************************************************************}
function PasDeDoublonChampCompte ( Table , Champ : string; Bourre : Char; Lg : integer; LCpt : TStringList; pszAxe : string = '' ) : boolean;
var T : TOB;
    Q : TQuery;
    i,j : integer;
    szSQL : string;
begin
  Result := True;
  if (TableToPrefixe(Table) = '') then exit;
  T := TOB.Create ('',nil,-1);
  try
    szSQL := 'SELECT * FROM '+Table;
    if pszAxe <> '' then szSQL := szSQL + ' WHERE S_AXE="'+pszAxe+'"';
    if Table = 'GENERAUX' then szSQL := szSQL + ' ORDER BY G_GENERAL'
    else if Table = 'TIERS' then szSQL := szSQL + ' ORDER BY T_AUXILIAIRE'
    else if Table = 'SECTION' then szSQL := szSQL + ' ORDER BY S_SECTION'
    else if Table = 'CSECTION' then szSQL := szSQL + ' ORDER BY CSP_SECTION';
    Q := OpenSQL (szSQL,True,-1,'',true);
    T.LoadDetailDB (Table,'','',Q,False);
    Ferme (Q);
    // Mise à jour des longueurs de compte
    for i:=0 to T.Detail.Count - 1 do
      T.Detail[i].PutValue(Champ,FormatCpte (T.Detail[i].GetValue(Champ), Bourre, Lg));
    // Contrôle des doublons
    for i:=0 to T.Detail.Count - 1 do
    begin
      for j:=i+1 to T.Detail.Count - 1 do
      begin
        if T.Detail[i].GetValue(Champ) = T.Detail[j].GetValue(Champ) then
        begin
          if not ParametrageCompteIdentique (T.Detail[i], T.Detail[j], Table) then
          begin
            LCpt.Add ( T.Detail[i].GetValue(Champ));
            Result := False;
            break;
          end else
          begin
            LCpt.Add ( T.Detail[i].GetValue(Champ));
          end;
        end;
      end;
      if not Result then break;
    end;
  finally
    T.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Changement de longueur de compte pour les champs
Suite ........ : compte d'une table quelconque
Mots clefs ... :
*****************************************************************}
procedure UpdateChampCompte ( Table : string ; AChamp : array of string; Bourre : Char; Lg : integer; pszType : String; PO : TInfoChgLgCpte; pszWhere : String ='');
var Prefixe, stBourrage, stSQL,stDateModif, stDPSTD : string;
    bDateModif : boolean;
    i, OldLg : integer;
begin
  if (TableToPrefixe(Table) = '') then exit;
  if assigned (PO) then PO ( 'Table '+Table+' - '+TableToLibelle( Table ));
  stBourrage := '';
  if pszType='AUX' then
    OldLg := VH^.Cpta[fbAux].Lg
  else
  if pszType='GEN' then
    OldLg := VH^.Cpta[fbGene].Lg
  else
    OldLg := VH^.Cpta[AxeToFB(pszType)].Lg;

  if Lg > OldLg then
    for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
  Prefixe := TableToPrefixe( Table );
  bDateModif := (ChampToNum(Prefixe+'_DATEMODIF')>0);
  if bDateModif then
    if DateWithHeure(Prefixe) then
      stDateModif := UsTime(NowH)
    else stDateModif := USDateTime(date) ;
  for i:=Low(AChamp)to High(AChamp) do
  begin
    stSQL := 'UPDATE '+Table+' SET ';
    if bDateModif then stSQL := stSQL + Prefixe+'_DATEMODIF="'+stDateModif+'",';
    stSQL := stSQL+AChamp[i]+'=';
    if OldLg > Lg then stSQL := stSQL + DB_MID(AChamp[i],1,Lg)
//    else stSQL := stSQL +AChamp[i]+'+"'+stBourrage+'"';
    else stSQL := stSQL +AChamp[i]+'||"'+stBourrage+'"';
    stSQL := stSQL + ' WHERE LEN('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere; // CA - 13/12/2002 - Pour compatibilité tte bases.
{    if V_PGI.Driver=dbINTRBASE then stSQL := stSQL + ' WHERE strlen('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;
    If V_PGI.Driver In [dbOracle7,dbOracle8] then stSQL := stSQL + ' WHERE length('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;
   	if V_PGI.Driver In [dbMSACCESS,dbMSSQL] then stSQL := stSQL + ' WHERE len('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;}
{$IFDEF EAGLCLIENT}
    stDPSTD := ''; // A FAIRE
{$ELSE}
    stDPSTD := SQLDPStd ( Table, False );
{$ENDIF}
    if stDPSTD <>'' then stSQL := stSQL + ' AND '+ stDPSTD;
    ExecuteSQL (stSQL);
{$IFDEF COMPTA}
    if ((Table='GENERAUX') or (Table='TIERS') or (Table='SECTION') or (Table='CSECTION')) then    // fiche 10301
        MAJHistoparam (Table,  '');
{$ENDIF}

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Changement de la longueur des comptes pour les champs
Suite ........ : compte de ParamSoc
Mots clefs ... :
*****************************************************************}
procedure UpdateCompteParamsoc ( AChamp : array of string ; Bourre : char; Lg : integer);
var Compte : string;
    i : integer;
begin
  for i:=Low(AChamp) to High (AChamp) do
  begin
    Compte := GetParamSocSecur ( AChamp[i], False);
    if Compte <> '' then
    begin
      Compte := FormatCpte (Compte, Bourre, Lg);
      SetParamSoc (AChamp[i],Compte);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. : - LG  - 28/111/2005 - ajout des comptes speciaux de
Suite ........ : revisions
Mots clefs ... :
*****************************************************************}
procedure ChangeLgCpteGen ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte);
begin
  // Domaine COMPTA  (C)
  UpdateChampCompte ('ANALYTIQ', ['Y_GENERAL','Y_CONTREPARTIEGEN'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('AXE', ['X_GENEATTENTE'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BANQUECP', ['BQ_GENERAL','BQ_COMPTEFRAIS'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDECR', ['BE_BUDGENE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDGENE', ['BG_BUDGENE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDJAL', ['BJ_GENEATTENTE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompteSpecif ( 'CBALSITECR',Bourre,Lg,'GEN',PO);
  UpdateChampCompte ('CORRESP', ['CR_CORRESP'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CPMODELEBQ', ['CMB_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CPMPTEMPOR', ['CTT_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CRELBQE', ['CRL_GENERAL','CRL_GENERALBQE','CRL_CONTREPARTIEGEN'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CRELBQEANALYTIQ', ['CRY_GENERAL','CRY_CONTREPARTIEGEN'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CROISCPT', ['CX_COMPTE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CROISEAXE', ['CCA_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CUMULDC', ['CQ_COMPTE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('DOCREGLE', ['DR_GENERAL','DR_CPTBANQUE','DR_BANQUEPREVI'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ECRGUI', ['EG_GENERAL','EG_BANQUEPREVI'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ECRITURE', ['E_GENERAL','E_BUDGET','E_CONTREPARTIEGEN','E_BANQUEPREVI'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('EEXBQ', ['EE_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('EEXBQLIG', ['CEL_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ETAPEREG', ['ER_CPTEDEPART','ER_CPTEARRIVEE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('FMTSUP', ['FS_GENATTEND','FS_CPTCOLLCLI','FS_CPTCOLLFOU'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('GENERAUX', ['G_GENERAL','G_CORRESP1','G_CORRESP2','G_BUDGENE','G_CODEIMPORT'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('HISTOBAL', ['HB_COMPTE1'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ICCECRITURE', ['ICE_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ICCEDTTEMP', ['ICZ_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ICCGENERAUX', ['ICG_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ICCTAUXCOMPTE', ['ICD_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMPECR', ['IE_GENERAL','IE_CONTREPARTIEGEN'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('JOURNAL', ['J_CONTREPARTIE','J_CPTEREGULDEBIT1','J_CPTEREGULDEBIT2','J_CPTEREGULDEBIT3','J_CPTEREGULCREDIT1','J_CPTEREGULCREDIT2','J_CPTEREGULCREDIT3'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('RBLIGECR', ['CRB_GENERAL'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('TXCPTTVA', ['TV_CPTEACH','TV_CPTEVTE','TV_ENCAISACH','TV_ENCAISVTE','TV_CPTACHRG','TV_CPTVTERG'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTIL', ['V_COMPTE'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('CPARAMGENER', ['CPG_GENERAL', 'CPG_ESCCPTGENE', 'CPG_ESCTVACPT', 'CPG_GENERALSEL'], Bourre, Lg, 'GEN', PO); // FQ 16842

  // GCO - 25/09/2007 - Révision Intégrée Compta
  UpdateChampCompte('CREVBLOCNOTE', ['CBN_CODE'], Bourre, Lg, 'GEN', PO, ' AND CBN_NATURE="GEN"');
  UpdateChampCompte('CREVGENERAUX', ['CRG_GENERAL'], Bourre, Lg, 'GEN', PO);
  // FIN GCO

  UpdateCompteParamsoc (['SO_BILDEB1','SO_BILDEB2','SO_BILDEB3','SO_BILDEB4','SO_BILDEB5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_BILFIN1','SO_BILFIN2','SO_BILFIN3','SO_BILFIN4','SO_BILFIN5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CHADEB1','SO_CHADEB2','SO_CHADEB3','SO_CHADEB4','SO_CHADEB5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CHAFIN1','SO_CHAFIN2','SO_CHAFIN3','SO_CHAFIN4','SO_CHAFIN5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEAMORTINF','SO_CPTEAMORTSUP','SO_CPTECBINF','SO_CPTECBSUP','SO_CPTEDEPOTINF'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEDEPOTSUP','SO_CPTEDEROGINF','SO_CPTEDEROGSUP','SO_CPTEDOTEXCINF','SO_CPTEDOTEXCSUP'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEDOTINF','SO_CPTEDOTSUP','SO_CPTEEXPLOITINF','SO_CPTEEXPLOITSUP','SO_CPTEFININF'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEFINSUP','SO_CPTEIMMOINF','SO_CPTEIMMOSUP','SO_CPTELOCINF','SO_CPTELOCSUP'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEPROVDERINF','SO_CPTEPROVDERSUP','SO_CPTEREPDERINF','SO_CPTEREPDERSUP','SO_CPTEREPEXCINF'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPTEREPEXCSUP','SO_CPTEVACEDEEINF','SO_CPTEVACEDEESUP','SO_DEFCOLCDIV','SO_DEFCOLCLI'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_DEFCOLDDIV','SO_DEFCOLDIV','SO_DEFCOLFOU'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_DEFCOLSAL','SO_GENATTEND'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_FERMEBEN','SO_FERMEBIL','SO_FERMEPERTE'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_OUVREBEN','SO_OUVREBIL','SO_OUVREPERTE'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_PRODEB1','SO_PRODEB2','SO_PRODEB3','SO_PRODEB4','SO_PRODEB5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_PROFIN1','SO_PROFIN2','SO_PROFIN3','SO_PROFIN4','SO_PROFIN5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_RESULTAT','SO_LETCHOIXGEN','SO_LETCHOIXGENC','SO_ICCCOMPTECAPITAL'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_CPPCA','SO_CPCCA'], Bourre, Lg);

// DOMAINE COMMUN (Y)
  UpdateChampCompte ('DEVISE', ['D_CPTLETTRDEBIT','D_CPTLETTRCREDIT','D_CPTPROVDEBIT','D_CPTPROVCREDIT'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('TIERS', ['T_COLLECTIF'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('TIERSCOMPL', ['YTC_SCHEMAGEN'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('MODEPAIE', ['MP_GENERAL','MP_CPTEREGLE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_BILDEB1','SO_BILDEB2','SO_BILDEB3','SO_BILDEB4','SO_BILDEB5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_BILFIN1','SO_BILFIN2','SO_BILFIN3','SO_BILFIN4','SO_BILFIN5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_CHADEB1','SO_CHADEB2','SO_CHADEB3','SO_CHADEB4','SO_CHADEB5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_CHAFIN1','SO_CHAFIN2','SO_CHAFIN3','SO_CHAFIN4','SO_CHAFIN5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_PRODEB1','SO_PRODEB2','SO_PRODEB3','SO_PRODEB4','SO_PRODEB5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_PROFIN1','SO_PROFIN2','SO_PROFIN3','SO_PROFIN4','SO_PROFIN5'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_EXTDEB1','SO_EXTDEB2','SO_EXTFIN1','SO_EXTFIN2','SO_GENATTEND'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_FERMEBIL','SO_OUVREBIL','SO_RESULTAT','SO_FERMEPERTE','SO_OUVREPERTE','SO_FERMEBEN','SO_OUVREBEN'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('SOCIETE', ['SO_DEFCOLCLI','SO_DEFCOLFOU','SO_DEFCOLSAL','SO_DEFCOLDDIV','SO_DEFCOLCDIV','SO_DEFCOLDIV'], Bourre, Lg, 'GEN', PO);
//  UpdateChampCompte ('SOCIETE', ['SO_ECCEURODEBIT','SO_ECCEUROCREDIT'], Bourre, Lg, 'GEN', PO);

// DOMAINE IMMO (I)
  UpdateChampCompte ('IMMO', ['I_COMPTEIMMO','I_COMPTEAMORT','I_COMPTEDOTATION','I_COMPTEDEROG','I_REPRISEDEROG',
                        'I_PROVISDEROG','I_DOTATIONEXC','I_COMPTELIE','I_COMPTELIEGAR','I_VAOACEDEE',
                        'I_COMPTEREF','I_REPEXPLOIT','I_REPEXCEP','I_AMORTCEDE','I_VACEDEE'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMMOAMOR', ['IA_COMPTEIMMO'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMMOCPTE', ['PC_COMPTEIMMO','PC_COMPTEAMORT','PC_COMPTEDOTATION','PC_COMPTEDEROG','PC_REPRISEDEROG','PC_DOTATIONEXC','PC_VACEDEE','PC_AMORTCEDE','PC_VOACEDE','PC_PROVISDEROG','PC_REPEXPLOIT','PC_REPEXCEP'] , Bourre, Lg, 'GEN', PO);

// DOMAINE SERVANTISSIMO (0)
  UpdateChampCompte ('IMOREF', ['IRF_COMPTEREF'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IAIDSAI', ['IAI_COMPTE','IAI_COMPTEREF'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMOVI1', ['IV1_COMPTEREF','IV1_COMPTEREFNEW'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMOAVANVIR', ['IAZ_COMPTEREF'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_01_105','IPE_01_131','IPE_01_145','IPE_01_281','IPE_01_404'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_01_44551','IPE_01_44562','IPE_01_44566','IPE_01_44571','IPE_01_462','IPE_01_612'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_01_613','IPE_01_615','IPE_01_671','IPE_01_6752','IPE_01_681','IPE_01_687','IPE_01_68725'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_01_771','IPE_01_7752','IPE_01_777','IPE_01_781','IPE_01_787','IPE_01_78725','IPE_01_78726','IPE_03_105','IPE_03_21','IPE_03_281'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_03_681','IPE_03_78726','IPE_04_105','IPE_04_131','IPE_04_1451','IPE_04_21','IPE_04_281','IPE_04_404','IPE_04_44562','IPE_04_44571'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_04_462','IPE_04_612','IPE_04_613','IPE_04_615','IPE_04_6752','IPE_04_681','IPE_04_687','IPE_04_68725','IPE_04_7752','IPE_04_777','IPE_04_781'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_04_787','IPE_04_78725','IPE_04_78726','IPE_05_105','IPE_05_21','IPE_05_281','IPE_05_681','IPE_05_78726','IPE_06_105','IPE_06_21'],Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IPARAMECR',['IPE_06_281','IPE_06_681','IPE_06_78726'],Bourre, Lg, 'GEN', PO);

// DOMAINE PAIE (P)
  UpdateChampCompte ('GUIDEECRPAIE', ['PGC_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ORGANISMEPAIE', ['POG_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTICOTPAIE', ['PVT_RACINE1','PVT_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTIORGPAIE', ['PVO_RACINE1','PVO_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTIREMPAIE', ['PVS_RACINE1','PVS_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateCompteParamsoc (['SO_PGCPTNETAPAYER'], Bourre, Lg);

  // Domaine Gescom (G)
  UpdateChampCompte ('CODECPTA', ['GCP_CPTEGENEACH','GCP_CPTEGENEVTE','GCP_CPTEGENESTOCK','GCP_CPTEGENEVARSTK',
                          'GCP_CPTEGENESCACH','GCP_CPTEGENESCVTE','GCP_CPTEGENREMACH','GCP_CPTEGENREMVTE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('PARPIECE', ['GPP_FAR_FAE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('TIERSPIECE', ['GTP_FAR_FAE'] , Bourre, Lg, 'GEN', PO);
  UpdateCompteParamsoc (['SO_GCCPTEESCACH','SO_GCCPTEESCVTE','SO_GCCPTEREMACH','SO_GCCPTEREMVTE',
        'SO_GCCPTEESCVTE','SO_GCCPTEHTACH','SO_GCCPTEHTVTE','SO_GCCPTEPORTACH',
        'SO_GCCPTEPORTVTE','SO_GCCPTERGVTE','SO_GCCPTESTOCK','SO_GCCPTEVARSTK',
        'SO_GCVRTINTERNE','SO_GCECARTCREDIT','SO_GCECARTDEBIT'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_AFCUTOFFCPTAAE','SO_AFCUTOFFCPTAAR','SO_AFCUTOFFCPTCCA','SO_AFCUTOFFCPTCHA',
                          'SO_AFCUTOFFCPTFAE','SO_AFCUTOFFCPTFAR','SO_AFCUTOFFCPTPCA','SO_AFCUTOFFCPTPRO'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_AFCUTOFFCPTAAE','SO_AFCUTOFFCPTAAR','SO_AFCUTOFFCPTCCA','SO_AFCUTOFFCPTCHA',
                          'SO_AFCUTOFFCPTFAE','SO_AFCUTOFFCPTFAR','SO_AFCUTOFFCPTPCA','SO_AFCUTOFFCPTPRO'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_AFCUTOFFTVAAAE','SO_AFCUTOFFTVAAAR','SO_AFCUTOFFTVAFAE','SO_AFCUTOFFTVAFAR'], Bourre, Lg);

  {25/07/06 : En tréso multi sociétés, il n'y a plus de notion de général sauf sur les comptes de contrepartie}
  if EstComptaTreso then begin
    UpdateChampCompte ('COURTSTERMES', ['TCT_CONTREPARTIETR'], Bourre, Lg, 'GEN', PO);
    UpdateChampCompte ('TRECRITURE'  , ['TE_CONTREPARTIETR'] , Bourre, Lg, 'GEN', PO);
    UpdateChampCompte ('TROPCVM'     , ['TOP_CONTREPARTIE']  , Bourre, Lg, 'GEN', PO);
    UpdateChampCompte ('FLUXTRESO'   , ['TFT_GENERAL']       , Bourre, Lg, 'GEN', PO);
    UpdateCompteParamsoc (['SO_COLCLIENT', 'SO_COLFOURNISSEUR', 'SO_COLSALARIE',
                           'SO_COLDIVERS', 'SO_TIERSDEB'      , 'SO_TIERSCRED'], Bourre, Lg);
    {02/01/07 : Si on n'a pas lancé le moulinage de BQ_CODE}
    if not GetParamSocSecur('SO_TRMOULINEBQCODE', True) then begin
      UpdateChampCompte ('CONDITIONDEC'    , ['TCN_GENERAL'] , Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('CONDITIONEQUI'   , ['TCE_GENERAL'] , Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('CONDITIONFINPLAC', ['TCF_GENERAL'] , Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('CONDITIONVAL'    , ['TCV_GENERAL'] , Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('TRVENTEOPCVM'    , ['TVE_GENERAL'] , Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('COURTSTERMES'    , ['TCT_COMPTETR'], Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('TRECRITURE'      , ['TE_GENERAL'  ], Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('TROPCVM'         , ['TOP_GENERAL' ], Bourre, Lg, 'GEN', PO);
      UpdateChampCompte ('EQUILIBRAGE'     , ['TEQ_SGENERAL' , 'TEQ_DGENERAL'], Bourre, Lg, 'GEN', PO);
    end;
  end;
end;

procedure ChangeLgCpteAux ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte );
begin
// DOMAINE COMPTA (C)
  UpdateChampCompte ('ANALYTIQ', ['Y_AUXILIAIRE','Y_CONTREPARTIEAUX'], Bourre, Lg,'AUX',PO);
  UpdateChampCompteSpecif ( 'CBALSITECR',Bourre,Lg,'AUX',PO);
  UpdateChampCompte ('CPMPTEMPOR', ['CTT_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('CRELBQE', ['CRL_GENERAL','CRL_CONTREPARTIEAUX'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('CRELBQEANALYTIQ', ['CRY_AUXILIAIRE','CRY_CONTREPARTIEAUX'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('DOCFACT', ['DF_CONTREPARTIEAUX'] , Bourre, Lg,'AUX' ,PO);
  UpdateChampCompte ('DOCREGLE', ['DR_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('ECRGUI', ['EG_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('ECRITURE', ['E_AUXILIAIRE','E_CONTREPARTIEAUX'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('FMTSUP', ['FS_FOUATTEND','FS_CLIATTEND','FS_SALATTEND','FS_DIVATTEND'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('HISTOBAL', ['HB_COMPTE2'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('IMPECR', ['IE_AUXILIAIRE','IE_CONTREPARTIEAUX'] , Bourre, Lg,'AUX',PO);
  UpdateCompteParamsoc (['SO_CLIATTEND','SO_FOUATTEND'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_SALATTEND','SO_DIVATTEND'], Bourre, Lg);

// DOMAINE COMMUN (Y)
  UpdateChampCompte ('TIERS', ['T_AUXILIAIRE',{'T_TIERS', CA - 29/11/2002 }'T_CORRESP1','T_CORRESP2','T_FACTURE','T_PAYEUR'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('RIB', ['R_AUXILIAIRE'], Bourre, Lg, 'AUX',PO);
  UpdateChampCompte ('SOCIETE', ['SO_CLIATTEND','SO_FOUATTEND'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('TIERSCOMPL', ['YTC_AUXILIAIRE','YTC_TIERSLIVRE'{,'YTC_TIERS' CA - 29/11/2002 }], Bourre, Lg,'AUX',PO);
      //mcd 17/05/05 champ auxi supprimé UpdateChampCompte ('ANNUAIRE', ['ANN_AUXILIAIRE'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('CONTACT', ['C_AUXILIAIRE'], Bourre, Lg,'AUX',PO);
  UpdateChampCompteSpecif ( 'LIENSOLE',Bourre,Lg,'AUX',PO);

// DOMAINE PAIE (P)
  UpdateChampCompte ('ETABCOMPL', ['ETB_AUXILIAIRE'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('PAIEENCOURS', ['PPU_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('SALARIES', ['PSA_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('VIREMENTS', ['PVI_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);

// DOMAINE IMMO (I)
  UpdateChampCompte ('IMMO', ['I_ORGANISMECB'], Bourre, Lg,'AUX',PO);

// DOMAINE SERVANTISSIMO (0)
  UpdateChampCompte ('IMOFAC', ['IFC_COMPTEFOUR'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('IMOREF', ['IRF_COMPTEFOUR'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('IMOREP', ['IRP_COMPTEFOUR'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('IMOSO1', ['IS1_COMPTEFOUR','IS1_COMPTEDEB'], Bourre, Lg,'AUX',PO);      

  // Domaine (A)
    UpdateChampCompte ('RESSOURCE', ['ARS_AUXILIAIRE'], Bourre, Lg,'AUX',PO);

  // Domaine (GC)
  UpdateCompteParamsoc (['SO_GCFOUCPTADIFF','SO_GCCLICPTADIFF','SO_GCFOUCPTADIFFPART','SO_GCCLICPTADIFFPART'], Bourre, Lg);

  // Domaine GRC (R)
  UpdateChampCompte ('ACTIONS', ['RAC_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('ACTIONSCHAINEES', ['RCH_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
{CRM_CRM10814_CD_DEB}
  UpdateChampCompte ('PERSPECTIVES', ['RPE_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('PERSPHISTO', ['RPH_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('PROSPECTCOMPL', ['RPC_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('PROSPECTS', ['RPR_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('ACTIONINTERVENANT', ['RAI_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompte ('PROJETS', ['RPJ_AUXILIAIRE'] , Bourre, Lg, 'AUX', PO);
  UpdateChampCompteSpecif ( 'RTINFOS003',Bourre,Lg,'AUX',PO);  
  UpdateChampCompteSpecif ( 'RTINFOS001',Bourre,Lg,'AUX',PO);
  UpdateChampCompteSpecif ( 'RTINFOS006',Bourre,Lg,'AUX',PO);
{CRM_CRM10814_CD_FIN}
end;

procedure ChangeLgCpteAna ( Lg : integer; Bourre : Char; pszAxe : String; PO : TInfoChgLgCpte );
begin
  UpdateChampCompte ('ANAGUI', ['AG_SECTION'] , Bourre, Lg, pszAxe, PO, ' AND AG_AXE="'+pszAxe+'"');
  UpdateChampCompte ('ANAGUIREF', ['AGR_SECTION'] , Bourre, Lg, pszAxe, PO, ' AND AGR_AXE="'+pszAxe+'"');
  UpdateChampCompte ('ANALYTIQ', ['Y_SECTION'] , Bourre, Lg, pszAxe, PO, ' AND Y_AXE="'+pszAxe+'"');
  UpdateChampCompte ('AXE', ['X_SECTIONATTENTE'] , Bourre, Lg, pszAxe, PO, ' AND X_AXE="'+pszAxe+'"');
  UpdateChampCompte ('AXEREF', ['XRE_SECTIONATTENTE'] , Bourre, Lg, pszAxe, PO, ' AND XRE_AXE="'+pszAxe+'"');
	UpdateChampCompte ('CRELBQEANALYTIQ', ['CRY_SECTION'] , Bourre, Lg, pszAxe, PO, ' AND CRY_AXE="'+pszAxe+'"');
  UpdateChampCompte ('HISTOANALPAIE', ['PHA_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND PHA_AXE="'+pszAxe+'"');
  UpdateChampCompte ('IECRANA', ['ILA_SECTION1','ILA_SECTION2','ILA_SECTION3','ILA_SECTION4','ILA_SECTION5'] , Bourre, Lg,pszAxe,PO, '');
  UpdateChampCompte ('INTERCOMPTA', ['GIC_SECTANA1','GIC_SECTANA2','GIC_SECTANA3','GIC_SECTANA4','GIC_SECTANA5'] , Bourre, Lg,pszAxe,PO, ' AND GIC_NATUREPIECEG LIKE "%'+pszAxe[2]+'"');
  UpdateChampCompte ('PAIEVENTIL', ['PAV_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND PAV_NATURE LIKE "%'+pszAxe[2]+'"');
  if VH^.LiaisonY2ViaShare then UpdateChampCompte ('CSECTION', ['CSP_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND CSP_AXE="'+pszAxe+'"');
  UpdateChampCompte ('SECTION', ['S_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND S_AXE="'+pszAxe+'"');
  UpdateChampCompte ('SECTIONREF', ['SRE_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND SRE_AXE="'+pszAxe+'"');
  UpdateChampCompte ('VENTANA', ['YVA_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND YVA_AXE="'+pszAxe+'"');
  UpdateChampCompte ('VENTIL', ['V_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND V_NATURE LIKE "%'+pszAxe[2]+'"');
  UpdateChampCompte ('VENTILREF', ['VR_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND VR_NATURE LIKE "%'+pszAxe[2]+'"');
  UpdateChampCompteSpecif ( 'CBALSITECR', Bourre, Lg, pszAxe, PO);
end;

function FormatCpte (Chaine : String; Bourre : Char; Lg : integer) : String;
var
  Indice, Longueur : Integer;
  St               : String;
begin
  { CA - 06/10/2003 - Ne me semble pas utile et peut provoquer des problèmes dans le cas de
  comptes tels que : A 1234 et A00000 qui seront vus de manière identique }
//  Chaine := ReadTokenPipe (Chaine, ' ');
  Chaine := Trim( Chaine );

  Longueur := Length (Chaine);
  if Longueur > 17 then
  begin
    FormatCpte := Chaine; exit;
  end;
  if (Lg <= Longueur) then
     FormatCpte := Copy (Chaine,0, Lg)
  else
  begin
    Indice:=Longueur;
    St := '';
    repeat
      St  := St + Bourre;
      inc (Indice);
    until (Indice >= Lg);
    FormatCpte := Chaine+St;
  end;
end;

procedure UpdateChampCompteSpecif ( Table : string; Bourre : Char ; Lg : integer; pszType : String; PO : TInfoChgLgCpte);
var Q : TQuery;
    i,OldLg : integer;
    stChamp, stBourrage, stSQL : string;
begin
  if (TableToPrefixe(Table) = '') then exit;
  
  if assigned (PO) then PO ( 'Table '+Table+' - '+TableToLibelle( Table ));

  if Table = 'CBALSITECR' then
  begin
    if pszType = 'AUX' then
    begin
      OldLg := VH^.Cpta[fbAux].Lg;
      Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_TYPECUM FROM CBALSIT WHERE BSI_TYPECUM="G/T" OR BSI_TYPECUM="TIE"',True,-1,'',true);
      while not Q.Eof do
      begin
        stBourrage := '';
        if Lg > OldLg then
          for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
        if Q.FindField('BSI_TYPECUM').AsString='G/T' then stChamp:='BSE_COMPTE2' else stChamp := 'BSE_COMPTE1';
        stSQL := 'UPDATE '+Table+' SET '+stChamp+'=';
        if OldLg > Lg then stSQL := stSQL + DB_MID(stChamp,1,Lg)
//        else stSQL := stSQL +stChamp+'+"'+stBourrage+'"';
        else stSQL := stSQL +stChamp+'||"'+stBourrage+'"';
        stSQL := stSQL + ' WHERE BSE_CODEBAL="'+Q.FindField('BSI_CODEBAL').AsString+'"';
        ExecuteSQL (stSQL);
        Q.Next;
      end;
      Ferme (Q);
    end else if pszType='GEN' then
    begin
      OldLg := VH^.Cpta[fbGene].Lg;
      Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_TYPECUM FROM CBALSIT WHERE BSI_TYPECUM="G/T" OR BSI_TYPECUM="GEN"',True,-1,'',true);
      while not Q.Eof do
      begin
        stBourrage := '';
        if Lg > OldLg then
          for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
        stSQL := 'UPDATE '+Table+' SET BSE_COMPTE1=';
        if OldLg > Lg then stSQL := stSQL + DB_MID('BSE_COMPTE1',1,Lg)
        else stSQL := stSQL +'BSE_COMPTE1 || "'+stBourrage+'"';
        stSQL := stSQL + ' WHERE BSE_CODEBAL="'+Q.FindField('BSI_CODEBAL').AsString+'"';
        ExecuteSQL (stSQL);
        Q.Next;
      end;
      Ferme (Q);
    end else // cas analytique
    begin
      OldLg := VH^.Cpta[AxeToFB(pszType)].Lg;
      Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_TYPECUM FROM CBALSIT WHERE BSI_TYPECUM="G/A" OR BSI_TYPECUM="ANA"',True,-1,'',true);
      while not Q.Eof do
      begin
        // Quel est le champ à mettre à jour ?
        if Q.FindField('BSI_TYPECUM').AsString='G/A' then stChamp:='BSE_COMPTE2' else stChamp := 'BSE_COMPTE1';
        // si ce n'est pas le bon axe, on passe à la balance suivante
        if Q.FindField(stChamp).AsString <> pszType then continue;
        stBourrage := '';
        if Lg > OldLg then
          for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
        stSQL := 'UPDATE '+Table+' SET '+stChamp+'=';
        if OldLg > Lg then stSQL := stSQL + DB_MID(stChamp,1,Lg)
//        else stSQL := stSQL +stChamp+'+"'+stBourrage+'"';
        else stSQL := stSQL +stChamp+'||"'+stBourrage+'"';
        stSQL := stSQL + ' WHERE BSE_CODEBAL="'+Q.FindField('BSI_CODEBAL').AsString+'"';
        ExecuteSQL (stSQL);
        Q.Next;
      end;
      Ferme (Q);
    end;
  end else if Table ='LIENSOLE' then
  begin
    OldLg := VH^.Cpta[fbAux].Lg;
    stSQL := 'UPDATE LIENSOLE SET LO_IDENTIFIANT=';
    if Lg > OldLg then
      for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
    if OldLg > Lg then stSQL := stSQL + DB_MID('LO_IDENTIFIANT',1,Lg)
    else stSQL := stSQL +'LO_IDENTIFIANT||"'+stBourrage+'"';
    stSQL := stSQL + ' WHERE LO_TABLEBLOB="T" AND LEN(LO_IDENTIFIANT)='+IntToStr(OldLg);
    ExecuteSQL(stSQL);
  end else if Table = 'ARTICLE' then
  begin
    OldLg := VH^.Cpta[fbGene].Lg;
    Q := OpenSQL ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_TYPEARTICLE="FI"',True,-1,'',true);
    while not Q.Eof do
    begin
      stBourrage := '';
      if Lg > OldLg then
        for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
      stSQL := 'UPDATE '+Table+' SET GA_DESIGNATION1=';
      if OldLg > Lg then stSQL := stSQL + DB_MID('GA_DESIGNATION1',1,Lg)
      else stSQL := stSQL +'GA_DESIGNATION11 || "'+stBourrage+'"';
      stSQL := stSQL + ' WHERE GA_ARTICLE="'+Q.FindField('GA_ARTICLE').AsString+'"';
      ExecuteSQL (stSQL);
      Q.Next;
    end;
    Ferme (Q);
{CRM_CRM10814_CD_DEB}
  end else if (Table ='RTINFOS001') or (Table ='RTINFOS003') or (Table ='RTINFOS006') then
  begin
    stBourrage := '';
    OldLg := VH^.Cpta[fbAux].Lg;
    if Lg > OldLg then
      for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
    if Table ='RTINFOS001' then
    begin
      stSQL := 'UPDATE RTINFOS001 SET RD1_DATEMODIF="' + USTIME(NowH) + '",';
      stSQL := stSQL + 'RD1_CLEDATA=';
      if OldLg > Lg then stSQL := stSQL + DB_MID('RD1_CLEDATA',1,Lg)
      else stSQL := stSQL + DB_MID('RD1_CLEDATA',1,OldLg) + '||"' + stBourrage+'"';
      stSQL := stSQL + '||' + DB_MID('RD1_CLEDATA',OldLg+1,25-OldLg);
      stSQL := stSQL + ' WHERE ' + DB_MID('RD1_CLEDATA',OldLg+1,1) + ' = ";"';
      ExecuteSQL(stSQL);
    end else if Table ='RTINFOS003' then
    begin
      stSQL := 'UPDATE RTINFOS003 SET RD3_DATEMODIF="' + USTIME(NowH) + '",';
      stSQL := stSQL + 'RD3_CLEDATA=';
      if OldLg > Lg then stSQL := stSQL + DB_MID('RD3_CLEDATA',1,Lg)
      else stSQL := stSQL +'RD3_CLEDATA||"'+stBourrage+'"';
      stSQL := stSQL + ' WHERE LEN(RD3_CLEDATA)='+IntToStr(OldLg);
      ExecuteSQL(stSQL);
    end else
    begin
      // clé T;Auxiliaire;No contact
      stSQL := 'UPDATE RTINFOS006 SET RD6_DATEMODIF="' + USTIME(NowH) + '",';
      stSQL := stSQL + 'RD6_CLEDATA=';
      if OldLg > Lg then stSQL := stSQL + DB_MID('RD6_CLEDATA',1,Lg+2)
      else stSQL := stSQL + DB_MID('RD6_CLEDATA',1,OldLg+2) + '||"' + stBourrage+'"';
      stSQL := stSQL + '||' + DB_MID('RD6_CLEDATA',OldLg+3,25-(OldLg+2));
      stSQL := stSQL + ' WHERE ' + DB_MID('RD6_CLEDATA',OldLg+3,1) + ' = ";"';
      ExecuteSQL(stSQL);
    end;
{CRM_CRM10814_CD_FIN}
  end;
end;

end.
