import 'reflect-metadata';

export const PLUGIN_EVENTS_METADATA_KEY = 'teams:plugin:events';
export type PluginEventName = 'error' | 'activity' | 'custom';
export type EventMetadata = {
  /**
   * the property name
   */
  readonly key: string;

  /**
   * the event name
   */
  readonly name: PluginEventName;
};

/**
 * add an event emitter to your plugin
 * via `@Event(...)`
 */
export function Event (name: PluginEventName) {
  return (target: any, propertyKey: string | ClassFieldDecoratorContext) => {
    if (typeof propertyKey === 'string') {
      const TargetType = target.constructor;
      const targetEventsMetadata: Array<EventMetadata> =
        Reflect.getOwnMetadata(PLUGIN_EVENTS_METADATA_KEY, TargetType) || [];

      targetEventsMetadata.push({
        key: propertyKey,
        name,
      });

      Reflect.defineMetadata(PLUGIN_EVENTS_METADATA_KEY, targetEventsMetadata, TargetType);
    } else if (typeof propertyKey === 'object' && propertyKey.kind === 'field') {
      // TC39 Stage 3 decorator
      const ctx = propertyKey as ClassFieldDecoratorContext;
      const key = String(ctx.name);
      ctx.addInitializer(function (this: any) {
        const TargetType = this.constructor;
        const targetEventsMetadata: Array<EventMetadata> =
          Reflect.getOwnMetadata(PLUGIN_EVENTS_METADATA_KEY, TargetType) || [];

        if (!targetEventsMetadata.some(m => m.key === key)) {
          targetEventsMetadata.push({ key, name });
          Reflect.defineMetadata(PLUGIN_EVENTS_METADATA_KEY, targetEventsMetadata, TargetType);
        }
      });
    }
  };
}
