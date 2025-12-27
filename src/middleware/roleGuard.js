const ROLE_HIERARCHY = {
  SPOC: ["SPOC", "CALENDAR_TEAM", "ADMIN"],
  CALENDAR_TEAM: ["CALENDAR_TEAM", "ADMIN"],
  ADMIN: ["ADMIN"],
};

export default function roleGuard(...allowedRoles) {
  return (req, res, next) => {
    if (!req.user) {
      return res.status(401).json({ message: "Unauthenticated" });
    }

    if (!allowedRoles.includes(req.user.role)) {
      return res.status(403).json({
        message: "Access denied for role: " + req.user.role,
      });
    }

    next();
  };
}
