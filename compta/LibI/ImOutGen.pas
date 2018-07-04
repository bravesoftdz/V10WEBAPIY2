{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/07/2003
Modifié le ... : 20/04/2006
Description .. : - CA - 18/07/2003 - Suppression des champs euro
Suite ........ : - BTY - 04/06 FQ 17629 - Calcul de la PMVAlue :
Suite ........ : 1. Erreur d'évaluation de la durée en linéaire
Suite ........ : 2. Choix de la règle de calcul de la PMValue
Suite......... :    - CT Tout en court terme
Suite......... :    - LT Tout en long terme
Suite......... :    - NOR Règle normale (comme avant)
Suite......... :    - RSD Règle normale en ignorant la durée (<=> NOR sans tenir compte de l'âge)
Suite......... : - MBO - 11/06 - fonction de récupération des comptes de subvention
Mots clefs ... :
*****************************************************************}
unit ImOutGen;

interface

uses HCtrls, ImEnt, (*SysUtils,*)HEnt1, HDB
   {$IFDEF eAGLCLient}
   ,UTob
   {$ELSE}
   {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
   {$ENDIF eAGLClient}

   {$IFDEF SERIE1}
   ;
   {$ELSE}
   ,UtilPGI ;
   {$ENDIF}

type TImPMValue = record
        PCT : double;
        PLT : double;
        MCT : double;
        MLT : double;
      end;

Type
  TCompteAss = record
    Immo : string;
    Amort : string;
    Dotation : string;
    Derog : string;
    RepriseDerog : string;
    DotationExcep : string;
    VaCedee : string;
    AmortCede : string;
    VoaCede : string;
    ProvisDerog : string;
    RepriseExcep : string;
    RepriseExploit : string;
  end;
//EPZ 28/11/00
{ A FAIRE
    CreanceCessionActif : string;
    ProduitCessionActif : string;
    TVACollectee : string;
}
//EPZ 28/11/00

  //ajout mbo 27.10.06 pour subvention d'équipement
  TCompteSBV = record
    Dotation : string;
    Reprise : string;
    QuotePart:string;
  end;

  TypeIntegration = (toDotation,toEcheance);
  TCumul = record
    Debit : double;
    Credit : double;
  end;
  TDefCompte = record
    Compte:string;
    Libelle:string;
    bGeneral : boolean;
  end;

function ControleCompteCharge (Compte , Nat : string) : boolean;
function ControleCompteImmo (Compte , Nat : string) : boolean;
function ControleOrganisme (Compte : String) : boolean;
function ControleCompteDepotGarantie (Compte : string) : boolean;
procedure RazSocieteImmo;
function CAssAmortissement (CAssImmo : string) : string;
function CAssDotation (CAssImmo : string) : string;
function CAssDerog (CAssImmo : string): string;
function CAssRepriseDerog (CAssImmo : string): string;
function CAssProvisDerog (CAssImmo : string): string;
function CAssDotationExc (CAssDotation : string) : string;
function CAssVaCedee (CAssImmo : string): string;
function CAssRepExploit (CAssDotation : string) : string;
function CAssRepExc (CAssDotationExc : string) : string;
function AjusteCompteImmo  (st : string) : string;
procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
function VerifCompteAssocie (Sender : TObject) : boolean;
function InsertOrUpdateLesComptesAssocies(Table,Prefix,CodeImmo:string;ComptesAssocies: TCompteAss;AvecInsert:boolean): boolean ;
procedure GetComptesAssociesParDefaut(var Arecord: TCompteAss; CompteImmo: string) ;
// ajout mbo 27.10.06
procedure GetComptesSBV(var Arecord: TCompteSBV; CodeImmo: string) ;
// BTY 04/06 FQ 17629 Ajout règle de calcul de la PMValue
//function CalculPMValue (dtOpe, dtAchat, dtAmort : TDateTime; dPValue, dCumAntEco, dDotCessEco : double;stMethodeEco : string) : TImPMValue;
function CalculPMValue (dtOpe, dtAchat, dtAmort : TDateTime; dPValue, dCumAntEco, dDotCessEco : double;
                        stMethodeEco, RegleCession : string) : TImPMValue;

//EPZ 28/11/00
{ A FAIRE
function CAssCreanceActif(CAssImmo : string): string;
function CAssProduitActif(CAssImmo : string): string;
function CAssTVACollectee(CAssImmo : string): string;
}
//EPZ 28/11/00


implementation
// --------Fonctions de travail sur les comptes ------------

// Controle l'existence et la validité du compte de charge
function ControleCompteCharge (Compte , Nat : string) : boolean;
begin
  ImBourreLaDoncSurLesComptes(Compte);//EPZ 30/10/00
  Result :=  Presence ( 'GENERAUX','G_GENERAL',Compte );
  if Result then
  begin
    if (Nat = 'CB')  and
       ((Compte >= VHImmo^.CpteCBInf) and (Compte <= VHImmo^.CpteCBSup)) then result := True
    else if (Nat = 'LOC') and ((Compte >= VHImmo^.CpteLocInf) and (Compte <= VHImmo^.CpteLocSup)) then Result := True
    else Result := false;
  end;
end;

// Controle l'existence et la validité du compte d'immobilisation
function ControleCompteImmo (Compte , Nat : string) : boolean;
var Q : TQuery;
begin
  Compte := ImBourreLaDoncSurLesComptes(Compte); //EPZ 30/10/00
  Q := OpenSQL ('SELECT G_GENERAL,G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+Compte+'"',True);
  // Result :=  Presence ( 'GENERAUX','G_GENERAL',Compte );
  Result := not Q.Eof;
  if Result then
  begin
//    if ((Nat = 'CB') or (Nat = 'PRO')) and
    if ((Nat = 'CB') or (Nat = 'PRO') or (Nat = 'LOC')) and  // Activation I_COMPTEIMMO pour TP
       ((Compte >= VHImmo^.CpteImmoInf) and (Compte <= VHImmo^.CpteImmoSup) and
       (Q.FindField ('G_NATUREGENE').AsString = 'IMO')) then result := True
    else
    if (Nat = 'FI') and
       ((Compte >= VHImmo^.CpteFinInf) and (Compte <= VHImmo^.CpteFinSup) and
       (Q.FindField ('G_NATUREGENE').AsString = 'IMO')) then result := True
    else Result := false;
  end;
  Ferme (Q);
end;

// Controle l'existence de l'organisme
function ControleOrganisme (Compte : string) : boolean;
var Q : TQuery;
begin
  //Compte := BourreLaDoncSurLesComptes(Compte); //EPZ 30/10/00
  Q := OpenSQL ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+Compte+'" AND T_NATUREAUXI="FOU"',True);
  Result := not Q.Eof;
  Ferme (Q);
end;

// Controle l'existence et la validité du compte de dépôt de garantie
function ControleCompteDepotGarantie (Compte : string) : boolean;
begin
  Result :=  Presence ( 'GENERAUX','G_GENERAL',Compte );
  if Result then
  begin
    Compte := ImBourreLaDoncSurLesComptes(Compte); //EPZ 30/10/00
    if ((Compte >= VHImmo^.CpteDepotInf) and (Compte <= VHImmo^.CpteDepotSup)) then result := True
    else result := false;
  end;
end;

// --------Fonctions de travail sur les montant ------------

procedure RazSocieteImmo;
begin
  ExecuteSQL ('DELETE FROM IMMO');
  ExecuteSQL ('DELETE FROM IMMOECHE');
  ExecuteSQL ('DELETE FROM IMMOCPTE');
  ExecuteSQL ('DELETE FROM IMMOLOG');
  ExecuteSQL ('DELETE FROM IMMOAMOR');
end;

function AjusteCompteImmo  (st : string) : string;
begin
  result := ImBourreLaDoncSurLesComptes(st);
end;

function CAssAmortissement (CAssImmo : string) : string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if Cimmo <> '' then
  begin
    Insert('8', CImmo, 2);
    Result := AjusteCompteImmo(CImmo);
  end  else Result := '';
end;

function CAssDotation (CAssImmo : string) : string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo <> '' then
  begin
//EPZ 30/10/00
//    if CImmo[2]='0' then Result := AjusteCompteImmo('68111000') else
//       Result := AjusteCompteImmo('68112000');
    if CImmo[2]='0' then Result := AjusteCompteImmo('681110') else
       Result := AjusteCompteImmo('681120');
//EPZ 30/10/00
  end  else Result := '';
end;

function CAssDerog (CAssImmo : string): string ;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CpteDerogInf);
end;

function CAssRepriseDerog (CAssImmo : string): string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CpteRepDerInf);
end;

function CAssProvisDerog (CAssImmo : string): string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CpteProvDerInf);
end;

function CAssDotationExc (CAssDotation : string) : string;
var CDotation : string;
begin
  CDotation := CAssDotation;
  if CDotation <> '' then
  begin
    CDotation[3] := '7';
    Result := AjusteCompteImmo(CDotation);
  end else Result := '';
end;

function CAssVaCedee(CAssImmo : string) : string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
//EPZ 30/10/00
//    Result := AjusteCompteImmo('67500000');
  Result := AjusteCompteImmo('675000');
//EPZ 30/10/00
end;

function CAssRepExploit (CAssDotation : string) : string;
var CDotation : string;
begin
  CDotation := CAssDotation;
  if CDotation <> '' then
  begin
    CDotation[1] := '7';
    Result := AjusteCompteImmo(CDotation);
  end else Result := '';
end;

function CAssRepExc (CAssDotationExc : string) : string;
var CDotationExc : string;
begin
  CDotationExc := CAssDotationExc;
  if CDotationExc <> '' then
  begin
    CDotationExc[1] := '7';
    Result := AjusteCompteImmo(CDotationExc);
  end else Result := '';
end;

//EPZ 28/11/00
{ A FAIRE
function CAssCreanceActif(CAssImmo : string): string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CpteCreanCessActInf);
end;

function CAssProduitActif(CAssImmo : string): string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CptePdtCessActInf);
end;

function CAssTVACollectee(CAssImmo : string): string;
var CImmo : string;
begin
  CImmo := CAssImmo;
  if CImmo = ''  then begin Result := '' ;exit;end;
  Result := AjusteCompteImmo(VHImmo^.CpteTVACollecteeInf);
end;
}
//EPZ 28/11/00

// BTY 04/06 FQ 17629
//function CalculPMValue (dtOpe, dtAchat, dtAmort : TDateTime; dPValue, dCumAntEco, dDotCessEco : double ; stMethodeEco : string) : TImPMValue;
function CalculPMValue (dtOpe, dtAchat, dtAmort:TDateTime; dPValue, dCumAntEco, dDotCessEco:double;
         stMethodeEco, RegleCession : string) : TImPMValue;
var PMValue : TImPMValue; CumDot : double; bCourtTerme : boolean;
begin
  if stMethodeEco='' then stMethodeEco:='NAM';
  CumDot := dCumAntEco+dDotCessEco;
  PMValue.PCT := 0; PMValue.PLT := 0; PMValue.MCT := 0; PMValue.MLT := 0;
{  if stMethodeEco = 'NAM' then
  begin
    bCourtTerme := (dtOpe) < PlusMois(dtAchat, 24);
    if bCourtTerme then
    begin
      if dPValue > 0 then PMValue.PCT := dPValue
      else PMValue.MCT := (-1)*dPValue;
    end else
    begin
      if dPValue > 0 then PMValue.PLT := dPValue
      else PMValue.MLT := (-1)*dPValue;
    end;
  end else
  begin
    if stMethodeEco = 'LIN' then bCourtTerme:=(dtOpe<PlusMois(dtAmort, 24))
                            else bCourtTerme:=(dtOpe<PlusMois(dtAchat, 24));
    if bCourtTerme then
    begin
      if dPValue > 0 then PMValue.PCT := dPValue
      else PMValue.MCT := (-1)*dPValue;
    end else
    begin
      if dPValue < 0 then PMValue.MCT := (-1)*dPValue
      else
      begin
        if dPValue > CumDot then
        begin
          PMValue.PCT := CumDot;
          PMValue.PLT := dPValue-CumDot;
        end else PMValue.PCT := dPValue;
      end;
    end;
  end; }
  // BTY 04/06 FQ 17629
  bCourtTerme := (dtOpe) < PlusMois(dtAchat, 24);
  // Tout dpValue se trouve en court terme
  if RegleCession = 'CT' then
     begin
     if dPValue > 0 then PMValue.PCT := dPValue
     else PMValue.MCT := (-1)*dPValue;
     end
  // Tout dpValue se trouve en long terme
  else if RegleCession = 'LT' then
     begin
     if dPValue > 0 then PMValue.PLT := dPValue
     else PMValue.MLT := (-1)*dPValue;
     end
  // Soit calcul normal, soit le normal en ignorant la durée
  else if (RegleCession = '') or (RegleCession = 'NOR') or (RegleCession = 'RSD') then
     begin
     if bCourtTerme and (not (RegleCession = 'RSD')) then
        begin
        if dPValue > 0 then PMValue.PCT := dPValue
        else PMValue.MCT := (-1)*dPValue;
        end
     else
        begin
        if stMethodeEco = 'NAM' then
           begin
           if dPValue > 0 then PMValue.PLT := dPValue
           else PMValue.MLT := (-1)*dPValue;
           end
        else
           begin
           if dPValue < 0 then PMValue.MCT := (-1)*dPValue
           else
              begin
              if dPValue > CumDot then
                 begin
                 PMValue.PCT := CumDot;
                 PMValue.PLT := dPValue-CumDot;
                 end
              else PMValue.PCT := dPValue;
              end;
           end;
        end;
     end;
  //
  result := PMValue;
end;

procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
var Q: TQuery; CompteImmo,Prefix: string;  DetruitQ: boolean ;
begin
  DetruitQ:=(Qimmo=nil) ;
  {$IFDEF SERIE1}
  Prefix:='IC' ;
  {$ELSE}
  Prefix:='PC' ;
  {$ENDIF}
  if QImmo<>nil then
    begin
    Q:=QImmo ;
    Prefix:='I' ;
    end
    else
      Q:=OpenSQL('SELECT * FROM IMMOCPTE WHERE '+Prefix+'_COMPTEIMMO="'+RefCompteImmo+'"',TRUE);
  try
    if (Q<>nil) and (not Q.Eof or (QImmo<>nil)) then
      begin
      ARecord.Immo           :=Q.FindField(Prefix+'_COMPTEIMMO').AsString;;
      ARecord.Amort          :=Q.FindField(Prefix+'_COMPTEAMORT').AsString;
      ARecord.Dotation       :=Q.FindField(Prefix+'_COMPTEDOTATION').AsString;
      ARecord.Derog          :=Q.FindField(Prefix+'_COMPTEDEROG').AsString;
      ARecord.RepriseDerog   :=Q.FindField(Prefix+'_REPRISEDEROG').AsString;
      ARecord.DotationExcep  :=Q.FindField(Prefix+'_DOTATIONEXC').AsString;
      ARecord.ProvisDerog    :=Q.FindField(Prefix+'_PROVISDEROG').AsString;
      ARecord.VaCedee        :=Q.FindField(Prefix+'_VACEDEE').AsString;
      ARecord.AmortCede      :=Q.FindField(Prefix+'_AMORTCEDE').AsString;
      ARecord.VoaCede        :=Q.FindField(Prefix+'_VAOACEDEE').AsString;
      ARecord.RepriseExcep   :=Q.FindField(Prefix+'_REPEXCEP').AsString;
      ARecord.RepriseExploit :=Q.FindField(Prefix+'_REPEXPLOIT').AsString;
      end
      else
      begin
      GetComptesAssociesParDefaut(ARecord,CompteImmo) ;
      end;
  finally
    if DetruitQ then Ferme (Q);
  end ;
end;


function VerifCompteAssocie (Sender : TObject) : boolean;
var stSup,stInf,TypeCompte : string;
begin
  TypeCompte := ExtractSuffixe(THDBEdit(Sender).Name);
  if TypeCompte = 'COMPTEIMMO' then
   begin stSup := VHImmo^.CpteImmoSup;stInf:=VHImmo^.CpteImmoInf; end
  else   if TypeCompte = 'COMPTEAMORT' then
     begin stSup := VHImmo^.CpteAmortSup;stInf:=VHImmo^.CpteAmortInf; end
  else   if TypeCompte = 'COMPTEDOTATION' then
     begin stSup := VHImmo^.CpteDotSup;stInf:=VHImmo^.CpteDotInf; end
  else   if TypeCompte = 'REPEXPLOIT' then
     begin stSup := VHImmo^.CpteExploitSup;stInf:=VHImmo^.CpteExploitInf; end
  else   if TypeCompte = 'DOTATIONEXC' then
     begin stSup := VHImmo^.CpteDotExcSup;stInf:=VHImmo^.CpteDotExcInf; end
  else   if TypeCompte = 'REPEXCEP' then
     begin stSup := VHImmo^.CpteRepExcSup;stInf:=VHImmo^.CpteRepExcInf; end
  else   if TypeCompte = 'VACEDEE' then
     begin stSup := VHImmo^.CpteVaCedeeSup;stInf:=VHImmo^.CpteVaCedeeInf; end
  else   if TypeCompte = 'AMORTCEDE' then
     begin stSup := VHImmo^.CpteAmortSup;stInf:=VHImmo^.CpteAmortInf; end
  else   if (TypeCompte = 'VOACEDE') or (TypeCompte = 'VAOACEDEE') OR (TypeCompte='VOACEDEE') then
     begin stSup := VHImmo^.CpteImmoSup;stInf:=VHImmo^.CpteImmoInf; end
  else   if TypeCompte = 'COMPTEDEROG' then
     begin stSup := VHImmo^.CpteDerogSup;stInf:=VHImmo^.CpteDerogInf; end
  else   if TypeCompte = 'PROVISDEROG' then
     begin stSup := VHImmo^.CpteProvDerSup;stInf:=VHImmo^.CpteProvDerInf; end
//EPZ 28/11/00
{ A FAIRE
  else   if TypeCompte = 'CREANCESURCESS' then
     begin stSup := VHImmo^.CpteCreanCessActSup;stInf:=VHImmo^.CpteCreanCessActInf; end
  else   if TypeCompte = 'PDTCESSACTIF' then
     begin stSup := VHImmo^.CptePdtCessActSup;stInf:=VHImmo^.CptePdtCessActInf; end
  else   if TypeCompte = 'TVACOLLECTEE' then
     begin stSup := VHImmo^.CpteTVACollecteeSup;stInf:=VHImmo^.CpteTVACollecteeInf; end
}
//EPZ 28/11/00
  else   if TypeCompte = 'REPRISEDEROG' then
     begin stSup := VHImmo^.CpteRepDerSup;stInf:=VHImmo^.CpteRepDerInf; end;
  {$IFDEF eAGLClient}
  result := ((THEdit(Sender).text<= stSup) and (THEdit(Sender).text>= stInf))
  {$ELSE}
  result := ((THDBEdit(Sender).Field.AsString <= stSup) and (THDBEdit(Sender).Field.AsString >= stInf))
  {$ENDIF eAGLClient}
end;

function InsertOrUpdateLesComptesAssocies(Table,Prefix,CodeImmo:string;ComptesAssocies: TCompteAss;AvecInsert:boolean): boolean ;
var VoACede,wWhere:string ;
begin
result:=true ;
if Table='IMMO' then begin VoACede:='VAOACEDEE' ; wWhere:=Prefix+'IMMO="'+CodeImmo+'"' ; end
                else begin VoACede:='VOACEDE' ; wWhere:=Prefix+'COMPTEIMMO="'+ComptesAssocies.Immo+'"' ; end ;

if ExisteSQL('SELECT * FROM '+Table+' WHERE  '+wWhere) then
  begin
  result:=(ExecuteSQL('UPDATE '+Table+' SET '+Prefix+'COMPTEAMORT="'+ComptesAssocies.Amort+
    '",'+Prefix+'COMPTEDOTATION="'+ComptesAssocies.Dotation+'",'+Prefix+'COMPTEDEROG="'+ComptesAssocies.Derog+
    '",'+Prefix+'REPRISEDEROG="'+ComptesAssocies.RepriseDerog+'",'+Prefix+'PROVISDEROG="'+ComptesAssocies.ProvisDerog+
    '",'+Prefix+'DOTATIONEXC="'+ComptesAssocies.DotationExcep+'",'+Prefix+'VACEDEE="'+ComptesAssocies.VaCedee+
    '",'+Prefix+'AMORTCEDE="'+ComptesAssocies.AmortCede+'",'+Prefix+'REPEXPLOIT="'+ComptesAssocies.RepriseExploit+
    '",'+Prefix+'REPEXCEP="'+ComptesAssocies.RepriseExcep+'",'+Prefix+VoACede+'="'+ComptesAssocies.VoaCede+
    '" WHERE '+wWhere)<>0);
  end
  else
  begin
    if AvecInsert and (Table<>'IMMO') then
      begin
      result:=(ExecuteSQL('INSERT INTO '+Table
      +'('+Prefix+'COMPTEAMORT,'+Prefix+'COMPTEDOTATION,'+Prefix+'COMPTEDEROG'+','+Prefix+'REPRISEDEROG,'
      +Prefix+'PROVISDEROG,'+Prefix+'DOTATIONEXC,'+Prefix+'VACEDEE,'+Prefix+'AMORTCEDE,'
      +Prefix+'REPEXPLOIT,'+Prefix+'REPEXCEP,'+VoACede+') '
      +'VALUES ("'+ComptesAssocies.Amort+'","'+ComptesAssocies.Dotation+'","'+ComptesAssocies.Derog
      +'","'+ComptesAssocies.RepriseDerog+'","'+ComptesAssocies.ProvisDerog+'","'+ComptesAssocies.DotationExcep
      +'","'+ComptesAssocies.VaCedee+'","'+ComptesAssocies.AmortCede+'",="'+ComptesAssocies.RepriseExploit+'","'+ComptesAssocies.RepriseExcep
      +'","'+ComptesAssocies.VoaCede
      +'" WHERE '+Prefix+'COMPTEIMMO="'+ComptesAssocies.Immo+'"')<>0);
      end
  end ;
end ;

procedure GetComptesAssociesParDefaut(var Arecord: TCompteAss; CompteImmo: string) ;
begin
ARecord.Immo           :=CompteImmo;
ARecord.Amort          :=CAssAmortissement (CompteImmo);
ARecord.Dotation       :=CAssDotation (CompteImmo);
ARecord.Derog          :=CAssDerog (CompteImmo);
ARecord.RepriseDerog   :=CAssRepriseDerog (CompteImmo);
ARecord.DotationExcep  :=CAssDotationExc (ARecord.Dotation);
ARecord.ProvisDerog    :=CAssProvisDerog (CompteImmo);
ARecord.VaCedee        :=CAssVaCedee(CompteImmo);
ARecord.AmortCede      :=ARecord.Amort;
ARecord.VoaCede        :=CompteImmo;
ARecord.RepriseExcep   :=CAssRepExc (ARecord.DotationExcep);
ARecord.RepriseExploit :=CAssRepExploit (ARecord.Dotation);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 27/10/2006
Modifié le ... :   /  /
Description .. : procedure alimentant l'enreg comptes SBV
Suite ........ : à partir du compte de dotation SBV, on reconstitue
Suite ........ : le compte de reprise et le compte de quote part
Mots clefs ... :
*****************************************************************}
procedure GetComptesSBV(var Arecord: TCompteSBV; CodeImmo: string) ;
var
CompteSbv : string;
Q : TQuery;
begin

  Q := OpenSQL ('select I_CPTSBVB from IMMO where I_IMMO = "' + CodeImmo + '"', true );
  CompteSBV := Q.FindField('I_CPTSBVB').AsString;
  Ferme(Q);

  ARecord.Dotation :=CompteSbv;

  if copy(CompteSBV, 1,3) = '138' then
  begin
     ARecord.Reprise := '139' + copy (CompteSBV, 4, length(CompteSBV)-3);
     ARecord.QuotePart := '777' + copy (CompteSBV, 4, length(CompteSBV)-3);
  end else
  begin
     ARecord.Reprise := '139' + copy (CompteSBV, 3, length(CompteSBV)-3);
     ARecord.QuotePart := '777' + copy (CompteSBV, 3, length(CompteSBV)-3);
  end;
end ;



end.
