 /*
 * Author: Tomas Takac
 * Article: https://www.codeproject.com/Articles/869059/Topological-sorting-in-Csharp
 * License: https://www.codeproject.com/info/cpol10.aspx
 */
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeDevPnP.Core.Extensions
{
    public static class TopologicalSort
    {
        private static Func<T, IEnumerable<T>> RemapDependencies<T, TKey>(IEnumerable<T> source, Func<T, IEnumerable<TKey>> getDependencies, Func<T, TKey> getKey)
        {
            var map = source.ToDictionary(getKey);
            return item =>
            {
                var dependencies = getDependencies(item);
                return dependencies != null
                    ? dependencies.Select(key => map[key])
                    : null;
            };
        }

        public static IList<T> Sort<T, TKey>(this IEnumerable<T> source, Func<T, IEnumerable<TKey>> getDependencies, Func<T, TKey> getKey, bool ignoreCycles = false)
        {
            ICollection<T> source2 = (source as ICollection<T>) ?? source.ToArray();
            return Sort<T>(source2, RemapDependencies(source2, getDependencies, getKey), null, ignoreCycles);
        }

        public static IList<T> Sort<T>(this IEnumerable<T> source, Func<T, IEnumerable<T>> getDependencies, IEqualityComparer<T> comparer = null, bool ignoreCycles = false)
        {
            var sorted = new List<T>();
            var visited = new Dictionary<T, bool>(comparer);

            foreach (var item in source)
            {
                Visit(item, getDependencies, sorted, visited, ignoreCycles);
            }

            return sorted;
        }

        private static void Visit<T>(T item, Func<T, IEnumerable<T>> getDependencies, List<T> sorted, Dictionary<T, bool> visited, bool ignoreCycles)
        {
            bool inProcess;
            var alreadyVisited = visited.TryGetValue(item, out inProcess);

            if (alreadyVisited)
            {
                if (inProcess && !ignoreCycles)
                {
                    throw new ArgumentException("Cyclic dependency found.");
                }
            }
            else
            {
                visited[item] = true;

                var dependencies = getDependencies(item);
                if (dependencies != null)
                {
                    foreach (var dependency in dependencies)
                    {
                        Visit(dependency, getDependencies, sorted, visited, ignoreCycles);
                    }
                }

                visited[item] = false;
                sorted.Add(item);
            }
        }

        public static IList<ICollection<T>> Group<T, TKey>(this IEnumerable<T> source, Func<T, IEnumerable<TKey>> getDependencies, Func<T, TKey> getKey, bool ignoreCycles = true)
        {
            ICollection<T> source2 = (source as ICollection<T>) ?? source.ToArray();
            return Group<T>(source2, RemapDependencies(source2, getDependencies, getKey), null, ignoreCycles);
        }

        public static IList<ICollection<T>> Group<T>(this IEnumerable<T> source, Func<T, IEnumerable<T>> getDependencies, IEqualityComparer<T> comparer = null, bool ignoreCycles = true)
        {
            var sorted = new List<ICollection<T>>();
            var visited = new Dictionary<T, int>(comparer);

            foreach (var item in source)
            {
                Visit(item, getDependencies, sorted, visited, ignoreCycles);
            }

            return sorted;
        }

        private static int Visit<T>(T item, Func<T, IEnumerable<T>> getDependencies, List<ICollection<T>> sorted, Dictionary<T, int> visited, bool ignoreCycles)
        {
            const int inProcess = -1;
            int level;
            var alreadyVisited = visited.TryGetValue(item, out level);

            if (alreadyVisited)
            {
                if (level == inProcess && ignoreCycles)
                {
                    throw new ArgumentException("Cyclic dependency found.");
                }
            }
            else
            {
                visited[item] = (level = inProcess);

                var dependencies = getDependencies(item);
                if (dependencies != null)
                {
                    foreach (var dependency in dependencies)
                    {
                        var depLevel = Visit(dependency, getDependencies, sorted, visited, ignoreCycles);
                        level = Math.Max(level, depLevel);
                    }
                }

                visited[item] = ++level;
                while (sorted.Count <= level)
                {
                    sorted.Add(new Collection<T>());
                }
                sorted[level].Add(item);
            }

            return level;
        }

    }
}
