export interface KardexRecord {
  id?: string;
  userName: string;
  folio: string;
  courseName: string;
  grade: string;
  section: string;
  date: string;
  curp?: string;
  sexo?: string;
  edad?: string;
  fechaNacimiento?: string;
  semestre?: string;
  aprobo?: string;
  folioConstancia?: string;
  numInterno?: string;
  nombrePreferencia?: string;
  tipoCurso?: string;
  instructor?: string;
  periodoImparticion?: string;
  searchKeywords?: string[];
  uploadedAt?: string;
}

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}
