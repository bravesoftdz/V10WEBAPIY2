unit WSNomadeDecl;


interface
uses InvokeRegistry,SysUtils;

type
  TWSUserParam = class (TRemotable )
  private
     fInternalID : WideString;
     fCodeuser : WideString;
     fpassword : WideString;
  public
     constructor create;
     destructor destroy; override;
  published
     property UniqueID : WideString read fInternalID write fInternalID;
     property User : WideString read fCodeuser write fCodeuser;
     property Password : WideString read fpassword write fpassword;
  end;
  (*
  TWSUser = class (TRemotable )
  private
     fInternalID : WideString;
     fRessource : WideString;
     fCodeuser : WideString;
     fDataBaseServer : WideString;
     fDatabase : WideString;
     fresult : Integer;
     fresultlib : WideString;
  public
     constructor create;
     destructor destroy; override;
  published
     property InternalID : WideString read fInternalID write fInternalID;
     property Ressource : WideString read fRessource write fRessource;
     property Codeuser : WideString read fCodeuser write fCodeuser;
     property DataBaseServer : WideString read fDataBaseServer write fDataBaseServer;
     property Database : WideString read fDatabase write fDatabase;
     property CodeErreur : Integer read fresult write fresult;
     property LibErreur : WideString read fresultlib write fresultlib;
  end;
  *)
  (*
  TWSAppel = class (TRemotable)
  public
    fCodeAppel : WideString;
    fEtatAppel : WideString;
    fPriorite : Integer;
    fDesignation : WideString;
    fDateAppel : WideString;
    fNomClient : WideString;
    fNomContact : WideString;
    fTelContact : WideString;
    fAdresse1 : WideString;
    fAdresse2 : WideString;
    fCodePostal : WideString;
    fVille : WideString;
  public
     constructor create;
     destructor destroy; override;
  published
    property CodeAppel : WideString read fCodeAppel write fCodeAppel;
  	property EtatAppel : WideString read fEtatAppel write fEtatAppel;
    property Priorite : Integer read fPriorite write fPriorite;
    property Designation : WideString read fDesignation write fDesignation;
    property DateAppel : WideString read fDateAppel write fDateAppel;
    property NomClient : WideString read fNomClient write fNomClient;
    property NomContact : WideString read fNomContact write fNomContact;
    property TelContact : WideString read fTelContact write fTelContact;
    property Adresse1 : WideString read fAdresse1 write fAdresse1;
    property Adresse2 : WideString  read fAdresse2 write fAdresse2;
    property CodePostal : WideString  read fCodePostal write fCodePostal;
    property Ville : WideString  read fVille write fVille;
  end;
  *)
 //  TWSAppels = WideString;
const
    SepChamp  = '$0000A';
    SepLigne =  '$0000D';

implementation


{ TWSAppel }

{ TWSUserParam }

constructor TWSUserParam.create;
begin

end;

destructor TWSUserParam.destroy;
begin

  inherited;
end;

end.
