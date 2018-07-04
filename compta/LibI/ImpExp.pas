unit ImpExp;

interface

uses
  (*Windows, *)HSysMenu, hmsgbox, Classes, Dialogs, Forms, Utob;
const
  FormatEnteteCegid = '%-70.70s%-3.3s%-3.3s';
  FormatStructureCegid = '%-3.3s%8.8s%-2.2s%-17.17s%-1.1s%-17.17s%-35.35s%-35.35s%-3.3s%8.8s%-1.1s%20.20s%-1.1s%8.8s%-3.3s%10.10s%-3.3s%20.20s%20.20s%-3.3s%2.2s%2.2s';
type
    TEnteteCegid = class
    SocieteEmettrice : string; //char 70
    TypeFormat : string; // char 3;
    FichierUtilise : string; // char 3
  end;

    TStructureCegid = class
    CodeJournal : string;	// char 3
    DatePiece : string;		// date 8
    TypePiece : string; 	// char 2
    CompteGene : string; 	// char 17
    TypeCompte : string;		// char 1
    AuxiliaireSection : string; 	// char 17
    RefInterne : string; 	// char 35
    Libelle : string; 	// char 35
    ModePaiement : string; 	// char 3
    Echeance : string;		// date 8
    Sens : string;		// char 1
    Montant1 : double;	// char 20
    TypeEcriture : string;	// char 1
    NumeroPiece : integer; // char 7
    Devise : string;	// char 3
    TauxDev : double;	// char 10
    CodeMontant : string;	// char 3
    Montant2 : double;	// char 20
    Montant3 : double;	// char 20
    Etablissement : string; // char 3
    Axe : string;	// char 2
    NumEche : integer; // char 2
  end;

  TFImportExport = class(TForm)
    SelectionFichier: TSaveDialog;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure AfficheCompteRenduExport(FichierExport : string);
function ChoixFichierExport(var FichierExport : string) : integer;
procedure InitEnteteCegid(var SEntete : TEnteteCegid; Societe, TypeFormat, FichierUtilise : string);
procedure InitStructureCegid(var SCegid : TStructureCegid);
procedure FormateStructureCegid(var LigneCegid : string ; SCegid : TStructureCegid);
procedure VentileStructureCegid(var SCegid : TStructureCegid; OB : TOB);
procedure FormateEnteteCegid(var LigneCegid : string ; SEntete : TEnteteCegid);
procedure CreeEnregIE(var SEntete : TEnteteCegid; var SCegid : TStructureCegid );
procedure DetruitEnregIE(var SEntete : TEnteteCegid; var SCegid : TStructureCegid );

implementation
(*, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HEnt1, ImEnt, HCtrls, hmsgbox, UiUtil, ImGenEcr, HSysMenu, utob*)
uses SysUtils, Hent1, iment ;

{$R *.DFM}

procedure InitEnteteCegid(var SEntete : TEnteteCegid; Societe, TypeFormat, FichierUtilise : string);
begin
  SEntete.SocieteEmettrice := Format('%s',[Societe]);
  SEntete.TypeFormat := Format('%s',[TypeFormat]);
  SEntete.FichierUtilise := Format('%s',[FichierUtilise]);
end;

procedure InitStructureCegid(var SCegid : TStructureCegid);
begin
  SCegid.CodeJournal := '   ';
  SCegid.DatePiece := '01011900';
  SCegid.TypePiece := '  ';
  SCegid.CompteGene := '                 ';
  SCegid.TypeCompte := ' ';
  SCegid.AuxiliaireSection := '                 ';
  SCegid.RefInterne := '                                   ';
  SCegid.Libelle := '                                   ';
  SCegid.ModePaiement := '   ';
  SCegid.Echeance := '01011900';
  SCegid.Sens := ' ';
  SCegid.Montant1 := 0.00;
  SCegid.TypeEcriture := ' ';
  SCegid.NumeroPiece := 0;
  SCegid.Devise := '   ';
  SCegid.TauxDev := 0.00;
  SCegid.CodeMontant := '   ';
  SCegid.Montant2 := 0.00;
  SCegid.Montant3 :=  0.00;
  SCegid.Etablissement := '';
  SCegid.Axe := '';
  SCegid.NumEche := 1;
end;

procedure FormateStructureCegid(var LigneCegid : string ; SCegid : TStructureCegid);
var sNumEche,sNumeroPiece,sMontant1,sMontant2,sMontant3,sTauxDev : string;
begin
  sMontant1 := FloatToStrF(SCegid.Montant1,ffFixed,20,V_PGI.OkDecV);
  sMontant2 := FloatToStrF(SCegid.Montant2,ffFixed,20,V_PGI.OkDecV);
  sMontant3 := FloatToStrF(SCegid.Montant3,ffFixed,20,V_PGI.OkDecV);
  sTauxDev := FloatToStrF(SCegid.TauxDev,ffFixed,10,V_PGI.OkDecV);
  sNumeroPiece := Format('%8.8d',[SCegid.NumeroPiece]);
  sNumEche := Format('%2.2d',[SCegid.NumEche]);
  LigneCegid := Format(FormatStructureCegid,[SCegid.CodeJournal,SCegid.DatePiece,
                        SCegid.TypePiece,SCegid.CompteGene,SCegid.TypeCompte,
                        SCegid.AuxiliaireSection,SCegid.RefInterne,
                        SCegid.Libelle,SCegid.ModePaiement,SCegid.Echeance,
                        SCegid.Sens,sMontant1,SCegid.TypeEcriture,sNumeroPiece,
                        SCegid.Devise,sTauxDev,SCegid.CodeMontant,sMontant2,sMontant3,
                        SCegid.Etablissement,SCegid.Axe,sNumEche]);
end;

procedure FormateEnteteCegid(var LigneCegid : string ; SEntete : TEnteteCegid);
begin
  LigneCegid := Format(FormatEnteteCegid,[SEntete.SocieteEmettrice,SEntete.TypeFormat,SEntete.FichierUtilise]);
end;

procedure VentileStructureCegid(var SCegid : TStructureCegid; OB : TOB);
var  Montant : double;
     jj,mm,aaaa : word;
     sPref : string;
begin
  if (OB.NomTable='ECRITURE') then sPref := 'E_' else sPref := 'Y_';
  SCegid.CodeJournal := OB.GetValue (sPref+'JOURNAL');
  DecodeDate(OB.GetValue(sPref+'DATECOMPTABLE'),aaaa,mm,jj);
  SCegid.DatePiece := Format('%2.2d%2.2d%4.4d',[jj,mm,aaaa]);
  SCegid.TypePiece := 'OD';
  SCegid.Devise := OB.GetValue(sPref+'DEVISE');
  if (VHImmo^.TenueEuro) then SCegid.CodeMontant := 'E--'
  else SCegid.CodeMontant := 'F--';
  SCegid.CodeMontant := Copy(SCegid.Devise,1,1)+'--';
  SCegid.CompteGene := OB.GetValue(sPref+'GENERAL');
  if (sPref = 'E_') then
  begin
    if (OB.GetValue('E_AUXILIAIRE') <> '') then
    begin
      SCegid.TypeCompte := 'X';
      SCegid.AuxiliaireSection := OB.GetValue('E_AUXILIAIRE');
    end else
    begin
      SCegid.TypeCompte := '';
    end;
  end else
  begin
    SCegid.TypeCompte := 'A';
    SCegid.AuxiliaireSection := OB.GetValue('Y_SECTION');
  end;
  SCegid.Libelle := OB.GetValue(sPref+'LIBELLE');
  if OB.GetValue(sPref+'DEBIT') <> 0.00 then
  begin
    Montant := OB.GetValue(sPref+'DEBIT');
    SCegid.Sens := 'D';
  end  else
  begin
    Montant := OB.GetValue(sPref+'CREDIT');
    SCegid.Sens := 'C';
  end;
  if (sPref='E_') then
  begin
    SCegid.ModePaiement := OB.GetValue('E_MODEPAIE');
    SCegid.Echeance := FormatDateTime('ddmmyyyy',OB.GetValue('E_DATEECHEANCE'));
  end;
  SCegid.TypeEcriture := OB.GetValue(sPref+'QUALIFPIECE');
  SCegid.Montant1 := Montant;
  SCegid.Etablissement := OB.GetValue(sPref+'ETABLISSEMENT');
  SCegid.NumeroPiece := OB.GetValue(sPref+'NUMEROPIECE');
  SCegid.TauxDev := OB.GetValue(sPref+'TAUXDEV');
  if (sPref='Y_') then SCegid.Axe := OB.GetValue ('Y_AXE');
end;

procedure CreeEnregIE(var SEntete : TEnteteCegid; var SCegid : TStructureCegid );
begin
  SCegid := TStructureCegid.Create;
  SEntete := TEnteteCegid.Create;
end;

procedure DetruitEnregIE(var SEntete : TEnteteCegid; var SCegid : TStructureCegid );
begin
  SCegid.Destroy;
  SEntete.Destroy;
end;

function ChoixFichierExport(var FichierExport : string) : integer;
var
  FImpExp : TFImportExport;
  ret : integer;
begin
  ret := 0;
  FImpExp := TFImportExport.Create(Application);
  try
    if FImpExp.SelectionFichier.Execute then
    begin
      FichierExport := FImpExp.SelectionFichier.FileName;
      ret:=FImpExp.SelectionFichier.FilterIndex;
    end;
    finally
      result:=ret;
  end;
  FImpExp.Free;
end;

procedure AfficheCompteRenduExport(FichierExport : string);
var
  FImpExp : TFImportExport;
  Texte : string;
begin
  FImpExp := TFImportExport.Create(Application);
  try
    Texte := Format(FImpExp.HM.Mess[1],[FichierExport]);
    FImpExp.HM.Execute(0,'Exportation des écritures',Texte);
  finally
    FImpExp.Free;
  end;
end;

end.
