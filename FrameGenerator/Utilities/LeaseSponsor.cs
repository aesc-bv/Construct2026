//using System;
//using System.Runtime.Remoting.Lifetime;

//namespace AESCConstruct25.Utilities
//{
//    /// <summary>
//    /// Grants effectively infinite lease renewals.
//    /// </summary>
//    public class LeaseSponsor : MarshalByRefObject, ISponsor
//    {
//        public TimeSpan Renewal(ILease lease)
//            => TimeSpan.MaxValue;  // infinite renewal

//        // Prevent the sponsor itself from expiring
//        public override object InitializeLifetimeService()
//            => null;
//    }

//    /// <summary>
//    /// Static helper to register the LeaseSponsor on any MarshalByRefObject.
//    /// </summary>
//    public static class SponsorRegistration
//    {
//        public static void Register(object obj)
//        {
//            if (obj is MarshalByRefObject mbr)
//            {
//                var lease = (ILease)mbr.GetLifetimeService();
//                lease?.Register(new LeaseSponsor());
//            }
//        }
//    }
//}