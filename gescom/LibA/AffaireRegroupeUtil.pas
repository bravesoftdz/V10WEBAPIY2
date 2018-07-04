unit AffaireRegroupeUtil;

interface

uses Paramsoc,HEnt1,UTob,HCtrls,AffaireUtil,EntGC,SysUtils,FactUtil,HMsgBox,
     FactGrp,UTOM,AGLInit,AffEcheanceUtil

{$IFDEF BTP}
	 ,CalcOleGenericBTP
{$ENDIF}

{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
      ;

Type T_TypeAffaire =  Set of (tIsAffRef,tIsSSAff,tIsAff,tIsComplete,tIsLight) ;

// Types d'affaires gérées en fox du contexte des defines ...
Function GereSousAffaire : Boolean;
Function GereAffCompleteEtLight : Boolean;
function GereAvenantAff : Boolean;
Procedure FindTypeAffaire (Var TypeAffaire :T_TypeAffaire ; IsAffaireRef, AffaireRef,Affaire,AffComplete : String);

// initialisation
Function AffCompleteParDefaut : Boolean;
// Chargement des sous affaires
Function ChargeTobSSAffaire (TobSSAffaire : TOB ; Affaire,Champs : String; AvecAffAffiche : Boolean): integer;
Function ChargeTobAffaireRef (TobAffaire : TOB ; AffaireRef,Champs : String; AvecAffAffiche : Boolean): Boolean;

// Gestion des avenants
Function MajAffRemplaceeParAvenant  (TobAff : TOB): Boolean;
Function GetCodeAffAvSuivant ( CodeAff : String) : string;
implementation


Function GereSousAffaire : Boolean;
BEGIN
if ctxScot in V_PGI.PGIContexte then Begin Result := False; Exit; END;
Result := GetParamSoc('SO_AFGERESSAFFAIRE');
END;

Function GereAvenantAff: Boolean;
BEGIN
if ctxScot in V_PGI.PGIContexte then Begin Result := False; Exit; END;
Result := GetParamSoc('SO_AFFGESTIONAVENANT');
END;

Function GereAffCompleteEtLight : Boolean;
BEGIN
Result := False;
// mcd 24/04/02 if (ctxTempo in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) or (VH_GC.GASeria) then
if (not(ctxScot in V_PGI.PGIContexte)) or (ctxBTP in V_PGI.PGIContexte) or (VH_GC.GASeria) then
   Begin Result := True; Exit; END;
END;

Function AffCompleteParDefaut : Boolean;
BEGIN
Result := True;
{$IFDEF BTP}
Result := GetParamSoc('SO_AFCOMPLETE')
{$ELSE}
//mcd 24/04/02if (ctxTempo in V_PGI.PGIContexte) or (VH_GC.GASeria) then
if (not(ctxScot in V_PGI.PGIContexte)) or (VH_GC.GASeria) then
   Result := GetParamSoc('SO_AFCOMPLETE')
else if ctxGCAff in V_PGI.PGIContexte then Result:=False;
{$ENDIF}
END;

Procedure FindTypeAffaire (Var TypeAffaire :T_TypeAffaire ; IsAffaireRef, AffaireRef,Affaire,AffComplete : String);
begin
// Traitement affaire / sous Affaire
if IsAffaireRef ='X' then TypeAffaire := [tIsAffRef] else
if (AffaireRef =Affaire) then TypeAffaire := [tIsAff] else
TypeAffaire := [tIsSSAff];
// Traitement  Aff. complète ou light
if AffComplete = 'X' then TypeAffaire := Typeaffaire + [tIsComplete]
                     else TypeAffaire := Typeaffaire + [tIsLight];
end;

Function ChargeTobSSAffaire (TobSSAffaire : TOB ; Affaire,Champs : String; AvecAffAffiche : Boolean): integer;
Var Q : Tquery;
    Table : string;
    i : integer;
    TobDet : TOB;
BEGIN
Result := 0;
if Affaire = '' then Exit;
if TobSSAffaire = Nil Then Exit;
if Champs ='' then BEGIN Champs :='*'; Table:='AFFAIRE'; END
              else Table :='extrait Affaire';
if TobSSAffaire.detail.count <> 0 then TobSSAffaire.ClearDetail;
Q := OPENSQL ('SELECT '+ Champs +' From AFFAIRE where AFF_AFFAIREREF="'+ Affaire+'" AND AFF_AFFAIRE<>"'+ Affaire +'"',True,-1,'',true);
if Not Q.EOF then
   BEGIN
   TobSSAffaire.LoadDetailDB(Table,'','',Q,False);
   Result := 1;
   if AvecAffAffiche then
      BEGIN
      For i:=0 to TobSSAffaire.Detail.count-1 do
         BEGIN
         TobDet := TobSSAffaire.Detail[i];
         if i=0 then TobDet.AddChampsup('AFFAIREAFFICHE',True);
         {$IFDEF BTP}
         TobDet.Putvalue('AFFAIREAFFICHE',BTPCodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' '));
         {$ELSE}
         TobDet.Putvalue('AFFAIREAFFICHE',CodeAffaireAffiche(TobDet.GetValue('AFF_AFFAIRE'),' '));
         {$ENDIF}
         END;
      END;
   END;
Ferme(Q);
END;

Function ChargeTobAffaireRef (TobAffaire : TOB ; AffaireRef,Champs : String; AvecAffAffiche : Boolean): Boolean;
Var Q : Tquery;
    Table : string;
BEGIN
Result := false;
if TobAffaire = Nil then Exit;
if AffaireRef = '' then Exit;
if Champs ='' then BEGIN Champs :='*'; Table:='AFFAIRE'; END
              else Table :='extrait Affaire';

if TobAffaire.Detail.Count <> 0 then TobAffaire.ClearDetail;

Q := OPENSQL ('SELECT '+ Champs +' From AFFAIRE where AFF_AFFAIRE="'+ AffaireRef+'"',True,-1,'',true);
if Not Q.EOF then
   BEGIN
   TobAffaire.SelectDB('',Q);
   Result := True;
   if AvecAffAffiche then
      BEGIN
      TobAffaire.AddChampsup('AFFAIREAFFICHE',True);
      {$IFDEF BTP}
      TobAffaire.Putvalue('AFFAIREAFFICHE',BTPCodeAffaireAffiche(TobAffaire.GetValue('AFF_AFFAIRE'),' '));
      {$ELSE}
      TobAffaire.Putvalue('AFFAIREAFFICHE',CodeAffaireAffiche(TobAffaire.GetValue('AFF_AFFAIRE'),' '));
      {$ENDIF}
      END;
   END;
Ferme(Q);
END;




// ************************* Gestion des avenants ******************************
Function MajAffRemplaceeParAvenant (TobAff: TOB): Boolean;
Var DateFinGener : TDateTime;
    Codeaffaire{,UpdateDatefin} : string;

Begin
Result := False;
if TobAff = Nil then Exit;
CodeAffaire := TobAff.GetValue('AFF_AFFAIRE');
if CodeAffaire ='' then Exit;

// Suppression des écheances
SupEcheancesAffaire (CodeAffaire, False,iDate1900);
DateFinGener:= AjusteDateGenerSurEch (CodeAffaire,False,False,iDate1900,TobAff.GetValue('AFF_DATEDEBGENER'));
TobAff.PutValue( 'AFF_ETATAFFAIRE', 'CLO');
TobAff.PutValue( 'AFF_DATEFINGENER', DateFinGener);
if (TobAff.GetValue('AFF_DATEFIN')=idate2099) then TobAff.PutValue( 'AFF_DATEFIN', V_PGI.DateEntree);
end;

Function GetNumAvenantSuivant ( CodeAvenant : String) : string;
Var Av : integer;
begin
Result := '01';
if CodeAvenant = '' then Exit;
if IsNumeric(CodeAvenant) then
   begin
   Av := StrToInt(CodeAvenant); Inc(Av); Result := Format('%*.*d',[2,2,Av]);
   end;
end;

Function GetCodeAffAvSuivant ( CodeAff : String) : string;
Var Av : string;
    bOK : Boolean;
begin
bOk := False;
Av := Copy(CodeAff,16,2);
while Not(bOk) do
   begin
   Av := GetNumAvenantSuivant(Av);
   Result := Copy(CodeAff,1,15)+Av;
   // PL le 03/07/01 on sécurise en sortant sur le dernier numéro d'avenant possible
   // attention, les avenants 99 vont s'écraser
   if not(ExisteAffaire (Result,'')) or (Result='99') then bOk := true;
   end;
end;




end.

