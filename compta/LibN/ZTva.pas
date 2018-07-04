unit ZTva;

interface

Uses StdCtrls, Controls, Classes,forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // VCL
     Dialogs,
     Windows,
     {$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     ZCompte,
     Ent1,
     ParamSoc, // pour le GetParamSocSecur
     // Lib
     ULibEcriture,
     // AGL
     UTOB ;

type

 TZTva = Class
  private
   FTOBTauxTVA  : TOB;       // tob des taux de tva par regime fiscal
   FZCompte : TZCompte;
   FStCompteTVA : string;
   FRdTaux : double;
  public
   procedure   InitializeObject;
   destructor  Destroy ; override ;

   procedure Load ( vStGeneral : string ; vBoLigneAuDebit : boolean);
   
   property ZCompte : TZCompte read FZCompte write FZCompte;

   property StCompteTVA : string read FStCompteTVA write FStCompteTVA;
   property RdTaux    : double read FRdTaux write FRdTaux;

 end;

implementation


destructor TZTva.Destroy ;
begin
 if assigned(FTOBTauxTVA)  then FTOBTauxTVA.Free;
 FTOBTauxTVA := nil;
end ;

procedure TZTva.InitializeObject;
var
 lQ : TQuery;
begin

  // chargement des taux de tva en fonction du regime fiscal du dossier
 FTOBTauxTVA    := TOB.Create('',nil,-1);
 lQ             := nil;
 try

 lQ := OpenSQL('select TV_CODETAUX ,TV_REGIME ,TV_TAUXACH ,TV_TAUXVTE ,TV_CPTEACH ,TV_CPTEVTE ' +
               ' FROM TXCPTTVA ' +
               ' where TV_TVAOUTPF = "TX1" ' +
               ' and TV_REGIME = "' + GetParamSocSecur('SO_REGIMEDEFAUT','') + '" ' ,
               true );

  FTOBTauxTVA.LoadDetailDB('TXCPTTVA','','',lQ,true);

 finally
  if assigned(lQ) then Ferme(lQ);
 end; // try

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. : Affection du taux de tva et du compte de tva en fonction du
Suite ........ : compte d'imputation et du sens du mouvement
Mots clefs ... :
*****************************************************************}
procedure TZTva.Load ( vStGeneral : string ; vBoLigneAuDebit : boolean);
var
 lTOB        : TOB;
 lStCodeTVA  : string;
 l : string;
begin

 RdTaux := 0;
 FStCompteTVA   := '';

 if ZCompte = nil then
  begin
   PGIInfo('La variable ZCompte n''est pas initilalisée','Erreur programmme');
   exit;
  end; // if

 if ZCompte.IsTvaAutorise(vStGeneral) then
  begin

   lStCodeTVA := ZCompte.GetCodeTVAPourUnCompte(vStGeneral);
   lTOB := FTOBTauxTVA.FindFirst(['TV_CODETAUX'],[lStCodeTVA],false);
   if assigned(lTOB) then
    begin

      if vBoLigneAuDebit then
       begin
        RdTaux := lTOB.GetValue('TV_TAUXACH');
        l := lTOB.GetValue('TV_CPTEACH');
        FStCompteTVA := '';
        FStCompteTVA := lTOB.GetValue('TV_CPTEACH');
       end
        else
         begin
          RdTaux := lTOB.GetValue('TV_TAUXVTE');
          l := lTOB.GetValue('TV_CPTEVTE');
          FStCompteTVA   := lTOB.GetValue('TV_CPTEVTE');
         end;

    end // if
     else
      begin
       RdTaux := 0;
       FStCompteTVA   := '';
      end
   end // if
  else
   begin
    RdTaux := 0;
    FStCompteTVA   := '';
   end;


end;



end.
