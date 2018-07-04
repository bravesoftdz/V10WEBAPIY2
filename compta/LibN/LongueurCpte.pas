unit LongueurCpte;
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  Forms, Dialogs,
  StdCtrls, Grids, Hctrls, Mask, ExtCtrls, ComCtrls,
  Buttons, Hspliter, Ent1, HCompte,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,HEnt1,UtilPGI,AGLInit,MajTable,
  hmsgbox, uTOB;

type
  TInfoChgLgCpte = procedure ( Msg : string ) of object;

function FormatCpte (Chaine : String; Bourre : Char; Lg : integer) : String;
procedure ChangeLgCpteGen ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte);
procedure ChangeLgCpteAux ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte );
procedure ChangeLgCpteAna ( Lg : integer; Bourre : Char; pszAxe : String; PO : TInfoChgLgCpte );
function PasDeDoublonChampCompte ( Table , Champ : string; Bourre : Char; Lg : integer; LCpt : TStringList ) : boolean;
procedure UpdateChampCompteSpecif ( Table : string; Bourre : Char ; Lg : integer; bAux : boolean; PO : TInfoChgLgCpte);
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
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc, Hstatus;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/04/2002
Modifié le ... :   /  /
Description .. : Compare les paramétrages de 2 comptes
Mots clefs ... :
*****************************************************************}
function ParametrageCompteIdentique (T1, T2 : TOB; Table : string ) : boolean;
begin
  if Table = 'GENERAUX' then begin
    Result := (T1.GetValue('G_NATUREGENE')= T2.GetValue('G_NATUREGENE'))
        and (T1.GetValue('G_COLLECTIF')= T2.GetValue('G_COLLECTIF'))
        and (T1.GetValue('G_FERME')= T2.GetValue('G_FERME'))
        and (T1.GetValue('G_LETTRABLE')= T2.GetValue('G_LETTRABLE'))
        and (T1.GetValue('G_POINTABLE')= T2.GetValue('G_POINTABLE'))
        and (T1.GetValue('G_VENTILABLE1')= T2.GetValue('G_VENTILABLE1'))
        and (T1.GetValue('G_VENTILABLE2')= T2.GetValue('G_VENTILABLE2'))
        and (T1.GetValue('G_VENTILABLE3')= T2.GetValue('G_VENTILABLE3'))
        and (T1.GetValue('G_VENTILABLE4')= T2.GetValue('G_VENTILABLE4'))
        and (T1.GetValue('G_VENTILABLE5')= T2.GetValue('G_VENTILABLE5'));
  	end
  else if Table = 'TIERS' then begin
    Result := (T1.GetValue('T_NATUREAUXI')= T2.GetValue('T_NATUREAUXI'))
        and (T1.GetValue('T_DEVISE')= T2.GetValue('T_DEVISE'))
        and (T1.GetValue('T_MULTIDEVISE')= T2.GetValue('T_MULTIDEVISE'))
        and (T1.GetValue('T_LETTRABLE')= T2.GetValue('T_LETTRABLE'))
        and (T1.GetValue('T_TIERS')= T2.GetValue('T_TIERS'))
        and (T1.GetValue('T_FERME')= T2.GetValue('T_FERME'));
		end
  else begin
  	Result := (T1.GetValue('S_SENS')= T2.GetValue('S_SENS'))
	      and (T1.GetValue('S_CORRESP1')= T2.GetValue('S_CORRESP1'))
        and (T1.GetValue('S_CORRESP2')= T2.GetValue('S_CORRESP2'))
        and (T1.GetValue('S_FERME')= T2.GetValue('S_FERME'))
        and (T1.GetValue('S_CONFIDENTIEL')= T2.GetValue('S_CONFIDENTIEL'))
        and (T1.GetValue('S_AXE')= T2.GetValue('S_AXE'));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : Contrôle la possibilité de doublon suite à une réduction de la
Suite ........ : longueur des comptes;
Suite ........ : Fonctionne pour les tables TIERS et GENERAUX
Mots clefs ... :
*****************************************************************}
function PasDeDoublonChampCompte ( Table , Champ : string; Bourre : Char; Lg : integer; LCpt : TStringList ) : boolean;
var T : TOB;
    Q : TQuery;
    i,j : integer;
begin
  Result := True;
  T := TOB.Create ('',nil,-1);
  try
    Q := OpenSQL ('SELECT * FROM '+Table,True);
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
var Prefixe, stBourrage, stSQL,stDateModif : string;
    bDateModif : boolean;
    i, OldLg : integer;
begin
  if assigned (PO) then PO ( 'Table '+Table+' - '+TableToLibelle( Table ));
  stBourrage := '';
  if pszType='AUX' then OldLg := VH^.Cpta[fbAux].Lg
  else if pszType='GEN' then OldLg := VH^.Cpta[fbGene].Lg
  else OldLg := VH^.Cpta[AxeToFB(pszType)].Lg;

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
		else stSQL := stSQL +AChamp[i]+'+"'+stBourrage+'"';

    if V_PGI.Driver=dbINTRBASE then stSQL := stSQL + ' WHERE strlen('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;
    If V_PGI.Driver In [dbOracle7,dbOracle8] then stSQL := stSQL + ' WHERE length('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;
    if V_PGI.Driver In [dbMSACCESS,dbMSSQL,dbMSSQL2005] then stSQL := stSQL + ' WHERE len('+AChamp[i]+')='+IntToStr(OldLg)+pszWhere;  //10/08/2006 YMO Ajout test SQL2005

    ExecuteSQL (stSQL);
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
    Compte := GetParamSoc ( AChamp[i]);
    if Compte <> '' then
    begin
      Compte := FormatCpte (Compte, Bourre, Lg);
      SetParamSoc (AChamp[i],Compte);
    end;
  end;
end;

procedure ChangeLgCpteGen ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte);
begin
  // Domaine COMPTA  (C)
  UpdateChampCompte ('ANALYTIQ', ['Y_GENERAL','Y_CONTREPARTIEGEN'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('AXE', ['X_GENEATTENTE'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BANQUECP', ['BQ_GENERAL','BQ_COMPTEFRAIS'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDECR', ['BE_BUDGENE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDGENE', ['BG_BUDGENE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('BUDJAL', ['BJ_GENEATTENTE'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompteSpecif ( 'CBALSITECR',Bourre,Lg,False,PO);
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
  UpdateCompteParamsoc (['SO_DEFCOLDDIV','SO_DEFCOLDIV','SO_DEFCOLFOU','SO_ECCEUROCREDIT'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_DEFCOLSAL','SO_GENATTEND'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_ECCEURODEBIT','SO_FERMEBEN','SO_FERMEBIL','SO_FERMEPERTE'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_OUVREBEN','SO_OUVREBIL','SO_OUVREPERTE'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_PRODEB1','SO_PRODEB2','SO_PRODEB3','SO_PRODEB4','SO_PRODEB5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_PROFIN1','SO_PROFIN2','SO_PROFIN3','SO_PROFIN4','SO_PROFIN5'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_RESULTAT','SO_LETCHOIXGEN','SO_LETCHOIXGENC','SO_ICCCOMPTECAPITAL'], Bourre, Lg);

// DOMAINE COMMUN (Y)
  UpdateChampCompte ('DEVISE', ['D_CPTLETTRDEBIT','D_CPTLETTRCREDIT','D_CPTPROVDEBIT','D_CPTPROVCREDIT'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('TIERS', ['T_COLLECTIF'], Bourre, Lg, 'GEN', PO);
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
  UpdateChampCompte ('SOCIETE', ['SO_ECCEURODEBIT','SO_ECCEUROCREDIT'], Bourre, Lg, 'GEN', PO);

// DOMAINE IMMO (I)
  UpdateChampCompte ('IMMO', ['I_COMPTEIMMO','I_COMPTEAMORT','I_COMPTEDOTATION','I_COMPTEDEROG','I_REPRISEDEROG','I_PROVISDEROG','I_DOTATIONEXC','I_COMPTELIE','I_COMPTELIEGAR','I_VAOACEDEE','I_COMPTEREF'], Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMMOAMOR', ['IA_COMPTEIMMO'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('IMMOCPTE', ['PC_COMPTEIMMO','PC_COMPTEAMORT','PC_COMPTEDOTATION','PC_COMPTEDEROG','PC_REPRISEDEROG','PC_DOTATIONEXC','PC_VACEDEE','PC_AMORTCEDE','PC_VOACEDE','PC_PROVISDEROG','PC_REPEXPLOIT','PC_REPEXCEP'] , Bourre, Lg, 'GEN', PO);

// DOMAINE PAIE (P)
  UpdateChampCompte ('GUIDEECRPAIE', ['PGC_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('ORGANISMEPAIE', ['POG_GENERAL'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTICOTPAIE', ['PVT_RACINE1','PVT_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTIORGPAIE', ['PVO_RACINE1','PVO_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateChampCompte ('VENTIREMPAIE', ['PVS_RACINE1','PVS_RACINE2'] , Bourre, Lg, 'GEN', PO);
  UpdateCompteParamsoc (['SO_PGCPTNETAPAYER'], Bourre, Lg);
end;

procedure ChangeLgCpteAux ( Lg : integer; Bourre : Char; PO : TInfoChgLgCpte );
begin
// DOMAINE COMPTA (C)
  UpdateChampCompte ('ANALYTIQ', ['Y_AUXILIAIRE','Y_CONTREPARTIEAUX'], Bourre, Lg,'AUX',PO);
  UpdateChampCompteSpecif ( 'CBALSITECR',Bourre,Lg,True,PO);
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
  UpdateCompteParamsoc (['SO_CLIATTEND','SO_FOUATTEND','SO_FOUATTEND'], Bourre, Lg);
  UpdateCompteParamsoc (['SO_SALATTEND','SO_DIVATTEND'], Bourre, Lg);

// DOMAINE COMMUN (Y)
  UpdateChampCompte ('TIERS', ['T_AUXILIAIRE','T_TIERS','T_CORRESP1','T_CORRESP2'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('RIB', ['R_AUXILIAIRE'], Bourre, Lg, 'AUX',PO);
  UpdateChampCompte ('SOCIETE', ['SO_CLIATTEND','SO_FOUATTEND'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('TIERSCOMPL', ['YTC_AUXILIAIRE','YTC_TIERS'], Bourre, Lg,'AUX',PO);

// DOMAINE PAIE (P)
  UpdateChampCompte ('ETABCOMPL', ['ETB_AUXILIAIRE'], Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('PAIEENCOURS', ['PPU_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('SALARIES', ['PSA_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
  UpdateChampCompte ('VIREMENTS', ['PVI_AUXILIAIRE'] , Bourre, Lg,'AUX',PO);
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
  UpdateChampCompte ('SECTION', ['S_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND S_AXE="'+pszAxe+'"');
  UpdateChampCompte ('SECTIONREF', ['SRE_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND SRE_AXE="'+pszAxe+'"');
  UpdateChampCompte ('VENTANA', ['YVA_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND YVA_AXE="'+pszAxe+'"');
  UpdateChampCompte ('VENTIL', ['V_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND V_NATURE LIKE "%'+pszAxe[2]+'"');
  UpdateChampCompte ('VENTILREF', ['VR_SECTION'] , Bourre, Lg,pszAxe,PO, ' AND VR_NATURE LIKE "%'+pszAxe[2]+'"');
end;

function FormatCpte (Chaine : String; Bourre : Char; Lg : integer) : String;
var
  Indice, Longueur : Integer;
  St               : String;
begin
  Chaine := ReadTokenPipe (Chaine, ' ');
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

procedure UpdateChampCompteSpecif ( Table : string; Bourre : Char ; Lg : integer; bAux : boolean; PO : TInfoChgLgCpte);
var Q : TQuery;
    i,OldLg : integer;
    stChamp, stBourrage, stSQL : string;
begin
  if assigned (PO) then PO ( 'Table '+Table+' - '+TableToLibelle( Table ));

  if Table = 'CBALSITECR' then
  begin
    if bAux then
    begin
      OldLg := VH^.Cpta[fbAux].Lg;
      Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_TYPECUM FROM CBALSIT WHERE BSI_TYPECUM="G/T" OR BSI_TYPECUM="TIE"',True);
      while not Q.Eof do
      begin
        stBourrage := '';
        if Lg > OldLg then
          for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
        if Q.FindField('BSI_TYPECUM').AsString='G/T' then stChamp:='BSE_COMPTE2' else stChamp := 'BSE_COMPTE1';
        stSQL := 'UPDATE '+Table+' SET '+stChamp+'=';
        if OldLg > Lg then stSQL := stSQL + DB_MID(stChamp,1,Lg)
        else stSQL := stSQL +stChamp+'"'+stBourrage+'"';
        stSQL := stSQL + ' WHERE BSE_CODEBAL="'+Q.FindField('BSI_CODEBAL').AsString+'"';
        ExecuteSQL (stSQL);
        Q.Next;
      end;
      Ferme (Q);
    end else
    begin
      OldLg := VH^.Cpta[fbGene].Lg;
      Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_TYPECUM FROM CBALSIT WHERE BSI_TYPECUM="G/T" OR BSI_TYPECUM="GEN"',True);
      while not Q.Eof do
      begin
        stBourrage := '';
        if Lg > OldLg then
          for i:=1 to (Lg-OldLg) do stBourrage := stBourrage+Bourre;
        stSQL := 'UPDATE '+Table+' SET BSE_COMPTE1=';
        if OldLg > Lg then stSQL := stSQL + DB_MID(Q.FindField('BSE_COMPTE1').AsString,1,Lg)
        else stSQL := stSQL +'BSE_COMPTE1"'+stBourrage+'"';
        stSQL := stSQL + ' WHERE BSE_CODEBAL="'+Q.FindField('BSI_CODEBAL').AsString+'"';
        ExecuteSQL (stSQL);
        Q.Next;
      end;
      Ferme (Q);
    end;
  end;
end;

end.
