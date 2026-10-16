/* =====================================================================
   CG HUBS - central settings, used by every app in this repository.

   THIS REPOSITORY IS PUBLIC. Put IDs and addresses here only - never
   passwords, client secrets, staff data or customer data.
   (Client IDs and the tenant ID are not secrets: every browser that
   opens a Hub already receives them. Access is controlled by Microsoft
   sign-in and SharePoint permissions, not by hiding these values.)
   ===================================================================== */

const CG_TENANT = {
  TENANT_ID: "ea9e105c-2df6-47ef-a8e3-b4aa1d788423",
  SP_HOST:   "conceptcompg.sharepoint.com"
};

/* ---------------------------------------------------------------------
   The apps. One entry per folder in this repository.

   status:
     "on"      - shown on the launcher for everyone
     "off"     - hidden from the launcher; its address shows "not available"
     "preview" - shown only to the emails in CG_PREVIEW (for testing)
   This only hides an app. It is NOT security - SharePoint permissions
   still decide who can read or change data.

   Every app starts as "off" and is switched to "preview" when it has
   been brought into CG_HUBS, then "on" at switch-over.
   --------------------------------------------------------------------- */
const CG_APPS = [
  { key: "training",   folder: "training/",   name: "CG Training HUB",     unit: "Concept Training Services",
    site: "/sites/training.admin", clientId: "ebbfa226-c073-4421-a02a-606f8341e4ae", status: "preview" },

  { key: "psp",        folder: "psp/",        name: "PSP HUB",             unit: "Payroll Services PNG",
    site: "/sites/PSP_Test",       clientId: "d5aef866-584f-4226-ab11-c941888fb519", status: "off" },

  { key: "csa",        folder: "csa/",        name: "CSA HUB",             unit: "Brian Hart Consulting",
    site: "/sites/Concept_CSA",    clientId: "06c6cde1-c805-4bdf-94c8-b36213b557e9", status: "preview" },

  { key: "sla",        folder: "sla/",        name: "SLA HUB",             unit: "Internal & Client Support",
    site: "/sites/Helpdesk",       clientId: "2489de0c-f15b-40d0-9468-a54a81bd4f34", status: "preview" },

  { key: "time",       folder: "time/",       name: "CG Time",             unit: "Time, attendance & payroll",
    site: "/sites/CG_Time",        clientId: "b8139965-9ca3-493e-85b4-95f799f8e8ba", status: "preview" },

  { key: "mobile",     folder: "mobile/",     name: "Concept Time Mobile", unit: "Log consulting & support time",
    site: "/sites/CG_Time",        clientId: "ec5454e0-c794-419b-9ec9-464e9917ef3c", status: "preview" },

  { key: "attendance", folder: "attendance/", name: "CG Time Attendance",  unit: "Clock in / out",
    site: "/sites/CG_Time",        clientId: "04c72756-916b-4fed-aa36-a6771ea4e8dc", status: "preview" }
];

/* ---------------------------------------------------------------------
   The launcher (the CG HUBS front page). It has its own small Entra app
   registration, "CG HUBS Launcher", with only User.Read - it needs to
   know who is signed in, not read any SharePoint data.
   --------------------------------------------------------------------- */
const CG_LAUNCHER = {
  clientId: "447f8eb4-833f-4a54-868b-19d0dd821124",
  scopes: ["User.Read"]
};

// Sign-in emails that can see apps set to "preview"
const CG_PREVIEW = ["mscully@concept.com.pg"];
