import { prisma } from '@/lib/db/prisma';
import { AuthenticatedUser } from './session';
import { WorkbookPermType } from '@prisma/client';

// ==============================================================================
// CENTRALIZED WORKBOOK POLICY EVALUATOR
// Evaluates ownership, system roles, and explicit granular workbook permissions
// Hierarchy & Precedence:
// - ROOT_ADMIN: Full system-wide access
// - OWNER: Implicit FULL access without needing a permissions row
// - ADMIN: Implicit VIEW and EXPORT access across workbooks
// - SHARED USER: Evaluates single effective permission from @@id([workbookId, userId])
// ==============================================================================

export interface PolicyResult {
  allowed: boolean;
  reason?: string;
  statusCode?: number;
}

export async function canAccessWorkbook(
  user: AuthenticatedUser,
  workbookId: string,
  requiredPermission: WorkbookPermType = 'VIEW'
): Promise<PolicyResult> {
  if (!user || !workbookId) {
    return { allowed: false, reason: 'Invalid user or workbook ID', statusCode: 400 };
  }

  // 1. ROOT_ADMIN has full system access
  if (user.isSystemRoot || user.roles.includes('ROOT_ADMIN')) {
    return { allowed: true };
  }

  // 2. Query Workbook existence and owner (Anti-IDOR)
  const workbook = await prisma.workbook.findUnique({
    where: { id: workbookId, isDeleted: false },
    select: { id: true, ownerId: true },
  });

  if (!workbook) {
    return { allowed: false, reason: 'Workbook tidak ditemukan atau telah dihapus.', statusCode: 404 };
  }

  // 3. OWNER has implicit FULL access
  if (workbook.ownerId === user.id) {
    return { allowed: true };
  }

  // 4. System ADMIN has system-level VIEW and EXPORT access
  if (user.roles.includes('ADMIN') && (requiredPermission === 'VIEW' || requiredPermission === 'EXPORT')) {
    return { allowed: true };
  }

  // 5. Shared User: Query exact single effective permission row
  const permRow = await prisma.workbookPermission.findUnique({
    where: {
      workbookId_userId: {
        workbookId,
        userId: user.id,
      },
    },
    select: { permission: true },
  });

  if (!permRow) {
    return { allowed: false, reason: 'Anda tidak memiliki izin akses ke workbook ini.', statusCode: 403 };
  }

  const effectivePerm = permRow.permission;

  // Precedence evaluation
  if (effectivePerm === 'FULL') {
    return { allowed: true };
  }

  if (requiredPermission === 'VIEW' && ['VIEW', 'EDIT', 'EXPORT', 'SHARE', 'FULL'].includes(effectivePerm)) {
    return { allowed: true };
  }

  if (requiredPermission === 'EDIT' && effectivePerm === 'EDIT') {
    return { allowed: true };
  }

  if (requiredPermission === 'EXPORT' && effectivePerm === 'EXPORT') {
    return { allowed: true };
  }

  if (requiredPermission === 'SHARE' && effectivePerm === 'SHARE') {
    return { allowed: true };
  }

  return {
    allowed: false,
    reason: `Izin '${effectivePerm}' tidak mencukupi untuk melakukan operasi '${requiredPermission}'.`,
    statusCode: 403,
  };
}
