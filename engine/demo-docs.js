/**
 * Clause Craft - demo documents (fictional parties, addresses and terms).
 * Used by the browser app ("Load Demo Lease") and the Word add-in ("Insert demo").
 * Both drafts deliberately contain one-sided and missing provisions so the review has something to find.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.ClauseCraftDemoDocs = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var lease = {
    id: "lease",
    title: "Office Lease - Suite 240 (demo)",
    paragraphs: [
      "OFFICE LEASE",
      "This Office Lease (this \"Lease\") is made as of October 1, 2026 by and between PACIFIC MERIDIAN OWNER, LLC, a Delaware limited liability company (\"Landlord\"), and HARBORLINE ADVISORY GROUP, INC., a California corporation (\"Tenant\").",
      "1. Premises. Landlord leases to Tenant, and Tenant leases from Landlord, Suite 240 (the \"Premises\") on the second floor of the building located at 8 Meridian Way, Irvine, California 92618 (the \"Building\"), containing approximately 3,950 rentable square feet.",
      "2. Term. The term of this Lease (the \"Term\") shall be sixty-two (62) months, commencing on December 1, 2026 (the \"Commencement Date\") and expiring on January 31, 2032 (the \"Expiration Date\"), unless sooner terminated as provided herein.",
      "3. Base Rent. Tenant shall pay Base Rent in monthly installments of $12,442.50 ($3.15 per rentable square foot per month), in advance on the first day of each calendar month, without notice, demand, offset or deduction. Base Rent for months two (2) through three (3) shall be abated.",
      "4. Rent Adjustments. On each anniversary of the Commencement Date, the monthly Base Rent shall increase by four percent (4%) over the Base Rent payable for the immediately preceding month.",
      "5. Operating Expenses. This is a triple net (NNN) lease. Commencing on the Commencement Date, Tenant shall pay, as Additional Rent, Tenant's Share (4.85%) of all Operating Expenses, real property taxes and insurance costs incurred by Landlord in owning, operating, managing and maintaining the Building and the common areas. Operating Expenses include, without limitation, the cost of capital improvements, capital repairs and capital replacements, and a management fee equal to five percent (5%) of gross revenues of the Building. Landlord's estimate for 2027 is $1.05 per rentable square foot per month.",
      "6. Security Deposit. Upon execution of this Lease, Tenant shall deposit with Landlord the sum of $25,000.00 as security for Tenant's performance. Landlord may apply all or any portion of the Security Deposit to cure any default, and Tenant shall restore the Security Deposit within ten (15) days after demand. Tenant waives the provisions of California Civil Code Section 1950.7.",
      "7. Use. Tenant shall use the Premises solely for general office purposes and for no other purpose.",
      "8. Condition of Premises. Tenant has inspected the Premises and accepts the Premises in its AS IS condition, with all faults, and Landlord shall have no obligation to make any improvements or alterations to the Premises.",
      "9. Repairs and Maintenance. Tenant shall, at its sole cost, maintain the Premises in good condition and repair, including the heating, ventilating and air conditioning (HVAC) systems serving the Premises, and Tenant shall replace the HVAC equipment when necessary in Landlord's judgment.",
      "10. Alterations. Tenant shall not make any alterations to the Premises without Landlord's prior written consent, which Landlord may withhold in its sole and absolute discretion.",
      "11. Assignment and Subletting. Tenant shall not assign this Lease or sublet all or any part of the Premises without Landlord's prior written consent, which Landlord may grant or withhold in its sole and absolute discretion. Upon any request for consent, Landlord may elect to recapture the Premises and terminate this Lease as to the space proposed to be transferred. Tenant shall pay to Landlord one hundred percent (100%) of any rent or other consideration received by Tenant in excess of the Rent payable under this Lease.",
      "12. Insurance. Tenant shall maintain commercial general liability insurance with limits of not less than $2,000,000 per occurrence, naming Landlord and its lender as additional insureds, and shall deliver certificates of insurance to Landlord within five (5) days after the date of this Lease.",
      "13. Indemnity. Tenant shall indemnify, defend and hold harmless Landlord and its members, managers, agents and lenders from and against any and all claims, losses, damages and liabilities arising from the use or occupancy of the Premises or the Building, whether or not caused by the negligence of Landlord. Landlord shall not be liable for any damage to Tenant's property from any cause whatsoever.",
      "14. Utilities and Services. Landlord shall not be liable for any interruption or failure of utilities or services to the Premises, and no such interruption shall entitle Tenant to any abatement of Rent.",
      "15. Damage or Destruction. If the Premises are damaged by fire or other casualty, Landlord may elect either to restore the Premises or to terminate this Lease. Tenant waives the provisions of California Civil Code Sections 1932(2) and 1933(4).",
      "16. Subordination. This Lease is subordinate to all existing and future mortgages and deeds of trust encumbering the Building. Tenant shall execute any subordination agreement requested by Landlord within five (5) days after request, and Tenant appoints Landlord as Tenant's attorney-in-fact to execute such documents if Tenant fails to do so.",
      "17. Estoppel Certificates. Tenant shall deliver an estoppel certificate in the form requested by Landlord within five (5) days after Landlord's request.",
      "18. Relocation. Landlord may relocate Tenant at any time during the Term to other space in the Building designated by Landlord upon thirty (30) days' notice.",
      "19. Holding Over. If Tenant remains in possession after the Expiration Date, Tenant shall pay Base Rent at two hundred percent (200%) of the Base Rent in effect immediately prior to such holdover, and Tenant shall be liable for all damages, including consequential damages, resulting from such holdover.",
      "20. Default. Each of the following is an Event of Default: (a) Tenant fails to pay any Rent when due and such failure continues for three (3) days, without notice; or (b) Tenant fails to perform any other obligation and such failure continues for ten (10) days after written notice.",
      "21. Late Charge. Any Rent not received within three (3) days after its due date shall incur a late charge equal to ten percent (10%) of the overdue amount.",
      "22. Landlord's Entry. Landlord and its agents may enter the Premises at any time to inspect, repair or show the Premises to prospective purchasers, lenders or tenants.",
      "23. Landlord's Lien. Tenant grants Landlord a security interest in all of Tenant's furniture, fixtures, equipment and other personal property located in the Premises to secure Tenant's obligations under this Lease.",
      "24. Parking. Tenant may use four (4) unreserved parking spaces per 1,000 rentable square feet at Landlord's then-prevailing rates.",
      "25. Rules and Regulations. Tenant shall comply with the rules and regulations of the Building, as Landlord may amend them from time to time.",
      "26. Guaranty. As a condition of this Lease, Tenant's principal, Jordan Ellis, shall personally guarantee all of Tenant's obligations under this Lease.",
      "27. Attorneys' Fees. Tenant shall reimburse Landlord for all reasonable attorneys' fees incurred by Landlord in enforcing this Lease.",
      "28. Brokers. Each party represents that it has dealt with no broker in connection with this Lease other than eXp Commercial, whose commission shall be paid by Landlord pursuant to a separate agreement.",
      "29. Waiver of Jury Trial. LANDLORD AND TENANT EACH WAIVE TRIAL BY JURY IN ANY ACTION ARISING OUT OF THIS LEASE.",
      "30. Notices. All notices under this Lease shall be in writing and delivered by personal delivery or nationally recognized overnight courier to the addresses set forth below the parties' signatures.",
      "31. Governing Law. This Lease shall be governed by the laws of the State of California.",
      "IN WITNESS WHEREOF, the parties have executed this Lease as of the date first written above.",
      "LANDLORD: PACIFIC MERIDIAN OWNER, LLC",
      "By: ____________________",
      "TENANT: HARBORLINE ADVISORY GROUP, INC.",
      "By: ____________________"
    ]
  };

  var psa = {
    id: "psa",
    title: "Purchase and Sale Agreement - 4120 Canyon Crest Road (demo)",
    paragraphs: [
      "PURCHASE AND SALE AGREEMENT AND JOINT ESCROW INSTRUCTIONS",
      "This Purchase and Sale Agreement (this \"Agreement\") is entered into as of October 5, 2026 by and between COBALT RIDGE HOLDINGS, LLC, a California limited liability company (\"Seller\"), and COASTLINE NET LEASE PARTNERS, LLC, a Delaware limited liability company (\"Buyer\").",
      "1. PROPERTY",
      "Seller agrees to sell and Buyer agrees to purchase the land and the single-tenant retail building located at 4120 Canyon Crest Road, Perris, California 92571 (the \"Property\"), together with Seller's interest in the existing lease with the tenant of the Property (the \"Lease\").",
      "2. PURCHASE PRICE AND DEPOSIT",
      "The purchase price for the Property is Four Million Two Hundred Thousand Dollars ($4,200,000.00) (the \"Purchase Price\"). Within three (3) business days after the Opening of Escrow, Buyer shall deposit One Hundred Thousand Dollars ($100,000.00) (the \"Deposit\") with Escrow Holder. The Deposit shall be non-refundable to Buyer except in the event of Seller's default.",
      "3. DUE DILIGENCE",
      "Buyer shall have a period of twenty-one (21) days after the Opening of Escrow (the \"Due Diligence Period\") to inspect the Property. Buyer shall not conduct any invasive testing of the Property without Seller's prior written consent.",
      "4. TITLE",
      "Buyer shall take title subject only to the Permitted Exceptions, which shall include all matters of record. Seller shall have no obligation to cure or remove any title matter.",
      "5. FINANCING",
      "Buyer's obligations under this Agreement are not contingent upon Buyer obtaining financing.",
      "6. CLOSING",
      "The Close of Escrow shall occur on or before the date that is thirty (60) days after the expiration of the Due Diligence Period (the \"Closing Date\").",
      "7. PRORATIONS",
      "Real property taxes, rents and operating expenses shall be prorated as of the Closing Date.",
      "8. REPRESENTATIONS AND WARRANTIES",
      "Seller represents and warrants to Buyer that, to Seller's actual knowledge, there are no pending condemnation proceedings affecting the Property. Seller's representations and warranties shall survive the Close of Escrow for a period of three (3) months.",
      "9. AS-IS SALE; RELEASE",
      "Buyer acknowledges that it is purchasing the Property AS-IS, WHERE-IS, WITH ALL FAULTS. Buyer releases Seller from any and all claims, known or unknown, arising from the physical or environmental condition of the Property, and Buyer waives the benefits of California Civil Code Section 1542.",
      "10. TENANT ESTOPPEL",
      "Seller shall use commercially reasonable efforts to obtain an estoppel certificate from the tenant under the Lease, but delivery of such estoppel certificate shall not be a condition to Buyer's obligations.",
      "11. DEFAULT; REMEDIES",
      "If Seller defaults under this Agreement, Buyer's sole and exclusive remedy shall be the return of the Deposit. If Buyer defaults, Seller shall retain the Deposit as liquidated damages.",
      "Buyer's Initials: ______  Seller's Initials: ______",
      "12. ASSIGNMENT",
      "Buyer shall not assign this Agreement without Seller's prior written consent.",
      "13. CASUALTY AND CONDEMNATION",
      "If, prior to the Close of Escrow, the Property is damaged by casualty and the cost of repair exceeds ten percent (10%) of the Purchase Price, Buyer may terminate this Agreement by written notice to Seller.",
      "14. BROKERS",
      "Each party represents that it has dealt with no broker other than eXp Commercial in connection with this transaction.",
      "15. GOVERNING LAW",
      "This Agreement shall be governed by the laws of the State of California.",
      "SELLER: COBALT RIDGE HOLDINGS, LLC",
      "By: ____________________",
      "BUYER: COASTLINE NET LEASE PARTNERS, LLC",
      "By: ____________________"
    ]
  };

  return { lease: lease, psa: psa, all: [lease, psa] };
});
