const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM","DATA_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  DATA_TEAM: ["DATA_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

export default function roleGuard(...allowedRoles) {
  return (req, res, next) => {
    if (!req.user || !req.user.role) {
      return res.status(401).json({ message: "Unauthenticated" });
    }

    const userRole = req.user.role;

    // 🔥 Check hierarchy
    const hasAccess = allowedRoles.some((requiredRole) =>
      ROLE_HIERARCHY[requiredRole]?.includes(userRole)
    );

    if (!hasAccess) {
      return res.status(403).json({
        message: `Access denied for role: ${userRole}`,
      });
    }

    next();
  };
}
