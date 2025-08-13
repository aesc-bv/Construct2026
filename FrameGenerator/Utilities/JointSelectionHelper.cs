using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class JointSelectionHelper
    {
        public static List<Component> GetSelectedComponents(Window window)
        {
            var sel = window?.ActiveContext?.Selection;
            var ordered = new List<Component>();
            if (sel == null || sel.Count == 0)
            {
                //Logger.Log("GetSelectedComponents: nothing selected.");
                return ordered;
            }

            var items = sel.ToList();
            //Logger.Log($"GetSelectedComponents: raw selection count = {items.Count}.");

            Component lastComp = null;
            for (int i = 0; i < items.Count; i++)
            {
                Component comp = null;

                // 1) direct component?
                if (items[i] is Component c)
                {
                    comp = c;
                    //Logger.Log($"  Item {i}: Component '{c.Name}'");
                }
                else
                {
                    // 2) reflectively look for a generic GetAncestor<T>()
                    var m = items[i].GetType()
                                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                                    .FirstOrDefault(mi => mi.Name == "GetAncestor"
                                                       && mi.IsGenericMethodDefinition
                                                       && mi.GetGenericArguments().Length == 1);
                    if (m != null)
                    {
                        try
                        {
                            var mg = m.MakeGenericMethod(typeof(Component));
                            comp = mg.Invoke(items[i], null) as Component;
                            //Logger.Log($"  Item {i}: {items[i].GetType().Name} → ancestor Component = '{comp?.Name ?? "null"}'");
                        }
                        catch (Exception)
                        {
                            //Logger.Log($"  Item {i}: reflection error: {ex.Message}");
                        }
                    }
                    else
                    {
                        //Logger.Log($"  Item {i}: Unrecognized type {items[i].GetType().Name}; no GetAncestor<>()");
                    }
                }

                if (comp == null)
                    continue;

                // stash the very last-clicked
                if (i == items.Count - 1)
                    lastComp = comp;
                else if (!ordered.Contains(comp))
                    ordered.Add(comp);
            }

            if (lastComp != null && !ordered.Contains(lastComp))
                ordered.Add(lastComp);

            //Logger.Log($"GetSelectedComponents: returning {ordered.Count} component(s).");
            return ordered;
        }
    }
}
