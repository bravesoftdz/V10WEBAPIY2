unit LibTiers;

interface

uses classes, uTOB, ParamSoc, HCtrls, SysUtils;

type  TTiers = class ( TOB )
        private
          procedure SetLibelle ( St : string );
        public
          constructor   Create ( T : TOB ); reintroduce ; overload ;
          destructor    Destroy ; override ;
          procedure     Charge ( Auxiliaire : string );
          procedure     InitNouveau ( General, Auxiliaire : string ; FmtSisco2 : boolean = False );
        published
          property Libelle : string write SetLibelle;
      end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  Hent1,
  Ent1;

{ TTiers }

procedure TTiers.Charge(Auxiliaire: string);
begin
  SelectDB ('"'+Auxiliaire+'"',nil);
end;

constructor TTiers.Create(T: TOB);
begin
  inherited Create('TIERS',T, - 1);
end;

destructor TTiers.Destroy;
begin
  inherited Destroy;
end;

procedure TTiers.InitNouveau(General, Auxiliaire: string; FmtSisco2 : boolean = False);
var Libelle : string;
    NatureGene : string;
begin
  Auxiliaire := BourreLaDonc(Auxiliaire, fbAux {fbGene !!!});
  if Length(Auxiliaire)<=3 then exit ;
  PutValue('T_AUXILIAIRE',Auxiliaire) ;
  PutValue('T_TIERS',Auxiliaire);
  Libelle := TraduireMemoire ('CREE PAR CEGID');
  PutValue('T_ABREGE',Copy(Libelle,1,17)) ;
  PutValue ('T_LIBELLE', Libelle );
  if General <> '' then
  begin
    NatureGene := GetColonneSQL('GENERAUX','G_NATUREGENE','G_GENERAL="'+General+'"');
    if NatureGene = 'COC' then PutValue('T_NATUREAUXI','CLI')
    else if NatureGene = 'COS' then PutValue('T_NATUREAUXI','SAL')
    else if NatureGene = 'COF' then PutValue('T_NATUREAUXI','FOU')
    else PutValue('T_NATUREAUXI','DIV');
    PutValue('T_COLLECTIF',General);
  end
  else
  begin
    PutValue ('T_NATUREAUXI','DIV');
    PutValue ('T_COLLECTIF',GetParamSocSecur('SO_DEFCOLDIV',''));
    if FmtSisco2 then
    begin
      if (Auxiliaire[1]='0') or (Auxiliaire[1]='F') then
      begin
        PutValue ('T_NATUREAUXI','FOU');
        PutValue ('T_COLLECTIF',GetParamSocSecur('SO_DEFCOLFOU',''));
      end else if (Auxiliaire[1]='9') or (Auxiliaire[1]='C') then
      begin
        PutValue ('T_NATUREAUXI','CLI');
        PutValue ('T_COLLECTIF',GetParamSocSecur('SO_DEFCOLCLI',''));
      end else
      begin
      end;
    end;
  end;
  PutValue('T_MODEREGLE',GetParamSocSecur ('SO_GCMODEREGLEDEFAUT',''));
  PutValue('T_REGIMETVA',GetParamSocSecur ('SO_REGIMEDEFAUT',''));
  PutValue('T_TVAENCAISSEMENT', GetParamSocSecur ('SO_CODETVADEFAUT',''));
  PutValue('T_LETTRABLE','X');
  PutValue('T_SOLDEPROGRESSIF','X');
  PutValue('T_DEVISE',V_PGI.DevisePivot);
  PutValue('T_CONFIDENTIEL','0');
end;

procedure TTiers.SetLibelle(St: string);
begin
  St := Trim (St);
  if St <> '' then
  begin
    PutValue('T_LIBELLE',Copy ( St, 1, 35 ) );
    PutValue('T_ABREGE',Copy ( St, 1, 17 ) );
  end else
  begin
    PutValue('T_LIBELLE',TraduireMemoire('CREE PAR CEGID'));
    PutValue('T_ABREGE',TraduireMemoire('CREE PAR CEGID'));
  end;
end;

end.
