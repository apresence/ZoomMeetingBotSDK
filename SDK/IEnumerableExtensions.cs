namespace ZoomMeetingBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public static class IEnumerableExtensions
    {
        private static readonly Random defaultRNG = new Random();

        /// <summary>
        /// Retrieves a random element from an enumerable using an optionally provided random number generator "rand".
        /// If "rand" is not provided, a default global RNG will be used.
        /// </summary>
        public static T RandomElement<T>(this IEnumerable<T> enumerable, Random rng = null)
        {
            int index;

            if (rng == null)
            {
                lock (defaultRNG)
                {
                    index = defaultRNG.Next(0, enumerable.Count());
                }
            }
            else
            {
                index = rng.Next(0, enumerable.Count());
            }

            return enumerable.ElementAt(index);
        }
    }
}
