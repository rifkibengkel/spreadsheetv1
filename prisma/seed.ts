import { PrismaClient, UserStatus } from "@prisma/client";
import { hashPassword } from "../src/lib/auth/argon2";

const prisma = new PrismaClient();

// ==============================================================================
// SEED CONSTANTS & RBAC MATRIX
// ==============================================================================

const PERMISSIONS = [
  // Users Management
  {
    code: "USER_VIEW",
    category: "Users",
    description: "Melihat daftar user dan profil pengguna",
  },
  {
    code: "USER_CREATE",
    category: "Users",
    description: "Membuat user baru secara langsung oleh admin",
  },
  {
    code: "USER_APPROVE",
    category: "Users",
    description: "Menyetujui pendaftaran user (PENDING -> ACTIVE)",
  },
  {
    code: "USER_EDIT",
    category: "Users",
    description: "Mengubah status dan informasi user",
  },
  {
    code: "USER_DISABLE",
    category: "Users",
    description: "Menonaktifkan akun user",
  },

  // RBAC Management
  {
    code: "ROLE_MANAGE",
    category: "RBAC",
    description: "Mengelola role dan penetapan permission",
  },

  // Workbooks Management
  {
    code: "WORKBOOK_VIEW",
    category: "Workbooks",
    description: "Membuka dan membaca workbook",
  },
  {
    code: "WORKBOOK_CREATE",
    category: "Workbooks",
    description: "Membuat dan mengunggah workbook baru",
  },
  {
    code: "WORKBOOK_EDIT",
    category: "Workbooks",
    description: "Mengubah dan menyimpan versi workbook",
  },
  {
    code: "WORKBOOK_DELETE",
    category: "Workbooks",
    description: "Menghapus workbook (soft delete)",
  },
  {
    code: "WORKBOOK_EXPORT",
    category: "Workbooks",
    description: "Mengekspor file XLSX atau CSV",
  },
  {
    code: "WORKBOOK_SHARE",
    category: "Workbooks",
    description: "Membagikan akses workbook ke user lain",
  },

  // Background Jobs
  {
    code: "JOB_VIEW",
    category: "Jobs",
    description: "Melihat riwayat dan status antrean background job",
  },
  {
    code: "JOB_CANCEL",
    category: "Jobs",
    description: "Membatalkan background job milik sendiri atau antrean",
  },

  // Audit Logs
  {
    code: "AUDIT_VIEW",
    category: "Audit",
    description: "Melihat audit logs aktivitas sistem",
  },

  // System Settings
  {
    code: "SYSTEM_SETTINGS",
    category: "System",
    description: "Mengonfigurasi parameter sistem enterprise",
  },
];

const ROLES = [
  {
    name: "ROOT_ADMIN",
    description: "Super Administrator Sistem dengan akses mutlak tak terbatas",
    isSystem: true,
  },
  {
    name: "ADMIN",
    description: "Administrator sistem pengelola user, workbook, dan audit log",
    isSystem: true,
  },
  {
    name: "USER",
    description: "Pengguna standar spreadsheet aplikasi",
    isSystem: true,
  },
];

const ADMIN_PERMISSIONS = [
  "USER_VIEW",
  "USER_CREATE",
  "USER_APPROVE",
  "USER_EDIT",
  "USER_DISABLE",
  "WORKBOOK_VIEW",
  "WORKBOOK_CREATE",
  "WORKBOOK_EDIT",
  "WORKBOOK_DELETE",
  "WORKBOOK_EXPORT",
  "WORKBOOK_SHARE",
  "JOB_VIEW",
  "JOB_CANCEL",
  "AUDIT_VIEW",
];

const USER_PERMISSIONS = [
  "WORKBOOK_VIEW",
  "WORKBOOK_CREATE",
  "WORKBOOK_EDIT",
  "WORKBOOK_EXPORT",
  "WORKBOOK_SHARE",
  "JOB_VIEW",
  "JOB_CANCEL",
];

// ==============================================================================
// MAIN SEED EXECUTION
// ==============================================================================

async function main() {
  console.log("🌱 Starting Enterprise Database Seed...");

  // 1. Seed Permissions
  console.log("📋 Seeding Permissions...");
  const permissionMap = new Map<string, string>();

  for (const perm of PERMISSIONS) {
    const record = await prisma.permission.upsert({
      where: { code: perm.code },
      create: perm,
      update: {
        description: perm.description,
        category: perm.category,
      },
    });
    permissionMap.set(perm.code, record.id);
  }
  console.log(`✓ ${PERMISSIONS.length} Permissions initialized`);

  // 2. Seed Roles
  console.log("🛡️ Seeding Roles...");
  const roleMap = new Map<string, string>();

  for (const role of ROLES) {
    const record = await prisma.role.upsert({
      where: { name: role.name },
      create: role,
      update: {
        description: role.description,
        isSystem: role.isSystem,
      },
    });
    roleMap.set(role.name, record.id);
  }
  console.log(`✓ ${ROLES.length} Roles initialized`);

  // 3. Assign Role Permissions
  console.log("🔗 Mapping Role Permissions...");

  // ROOT_ADMIN -> ALL Permissions
  const rootRoleId = roleMap.get("ROOT_ADMIN")!;
  for (const [_permCode, permId] of permissionMap.entries()) {
    await prisma.rolePermission.upsert({
      where: {
        roleId_permissionId: { roleId: rootRoleId, permissionId: permId },
      },
      create: { roleId: rootRoleId, permissionId: permId },
      update: {},
    });
  }

  // ADMIN -> Admin Permissions
  const adminRoleId = roleMap.get("ADMIN")!;
  for (const permCode of ADMIN_PERMISSIONS) {
    const permId = permissionMap.get(permCode);
    if (permId) {
      await prisma.rolePermission.upsert({
        where: {
          roleId_permissionId: { roleId: adminRoleId, permissionId: permId },
        },
        create: { roleId: adminRoleId, permissionId: permId },
        update: {},
      });
    }
  }

  // USER -> User Permissions
  const userRoleId = roleMap.get("USER")!;
  for (const permCode of USER_PERMISSIONS) {
    const permId = permissionMap.get(permCode);
    if (permId) {
      await prisma.rolePermission.upsert({
        where: {
          roleId_permissionId: { roleId: userRoleId, permissionId: permId },
        },
        create: { roleId: userRoleId, permissionId: permId },
        update: {},
      });
    }
  }
  console.log("✓ Role Permissions mapped successfully");

  // 4. Deterministic Root Administrator Bootstrap
  console.log("👑 Checking Root Administrator invariant...");

  const rootUsername = process.env.INITIAL_ROOT_ADMIN_USERNAME || "rootadmin";
  const rootEmail =
    process.env.INITIAL_ROOT_ADMIN_EMAIL || "root@redbox.digital";
  const rootPassword =
    process.env.INITIAL_ROOT_ADMIN_PASSWORD || "RootAdminSecurePassword123!";
  const rootName =
    process.env.INITIAL_ROOT_ADMIN_NAME || "System Root Administrator";

  const existingRoot = await prisma.user.findFirst({
    where: {
      OR: [
        { isSystemRoot: true },
        { username: rootUsername },
        { email: rootEmail },
      ],
    },
  });

  if (existingRoot) {
    console.log(
      `✓ Root Administrator already exists (ID: ${existingRoot.id}, username: ${existingRoot.username}). Idempotent skip.`,
    );
  } else {
    console.log(
      `Creating deterministic Root Administrator (${rootUsername})...`,
    );
    const passwordHash = await hashPassword(rootPassword);

    const newRoot = await prisma.user.create({
      data: {
        username: rootUsername,
        email: rootEmail,
        fullName: rootName,
        passwordHash,
        status: UserStatus.ACTIVE,
        isSystemRoot: true,
      },
    });

    // Assign ROOT_ADMIN role
    await prisma.userRole.create({
      data: {
        userId: newRoot.id,
        roleId: rootRoleId,
      },
    });

    // Record initial bootstrap in Audit Log
    await prisma.auditLog.create({
      data: {
        userId: newRoot.id,
        action: "SYSTEM_BOOTSTRAP_ROOT_ADMIN",
        resourceType: "User",
        resourceId: newRoot.id,
        metadata: {
          username: rootUsername,
          email: rootEmail,
          note: "Deterministic bootstrap of system root administrator during seed",
        },
      },
    });

    console.log(
      `✓ Root Administrator successfully created (ID: ${newRoot.id}) with ROOT_ADMIN role!`,
    );
  }

  console.log("🎉 Database seeding completed with 0 errors.");
}

main()
  .catch((e) => {
    console.error("❌ Seeding failed:", e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
