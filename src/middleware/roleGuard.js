const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

export  const roleGuard = (...allowedRoles) => {
  return (req, res, next) => {
    const userRole = req.user.role;

    const allowed = allowedRoles.some(role =>
      ROLE_HIERARCHY[role]?.includes(userRole)
    );

    if (!allowed) {
      return res.status(403).json({ message: "Forbidden" });
    }

    next();
  };
};
